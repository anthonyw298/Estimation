/**
 * ML Cost Predictor - Browser-compatible linear regression
 * Matches Python's utils/ml_predictor.py behavior:
 *   4 features: [sqft, bays, width, height]
 *   3-tier prediction: exact match -> linear regression -> avg/sqft fallback
 *   Training data persisted to Supabase
 */

import { db } from './database';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface TrainingSample {
  id: string;
  project: string;
  elevation: string;
  width: number;
  height: number;
  bays: number;
  bays_wide: number;
  bays_tall: number;
  sqft: number;
  finish: string;
  cost: number;
  timestamp: string;
}

export interface PredictionResult {
  cost: number;
  confidence: number;
  method: 'exact_match' | 'ml_model' | 'avg_per_sqft' | 'no_data';
  training_samples: number;
}

export interface MLStatistics {
  sample_count: number;
  avg_cost: number;
  min_cost: number;
  max_cost: number;
  avg_width: number;
  avg_height: number;
  avg_sqft: number;
  avg_cost_per_sqft: number;
  common_configurations: Array<{
    config: string;
    count: number;
    percentage: number;
  }>;
}

export interface MLStatus {
  is_trained: boolean;
  sample_count: number;
  ml_available: boolean;
}

// ---------------------------------------------------------------------------
// Simple Linear Regression (no external deps)
// ---------------------------------------------------------------------------

class SimpleLinearRegression {
  private weights: number[] | null = null;
  private bias: number = 0;
  private featureMeans: number[] = [];
  private featureStds: number[] = [];
  private trained = false;

  /**
   * Train using ordinary least squares with feature standardization.
   * X: array of feature vectors [sqft, bays, width, height]
   * y: array of target values (cost)
   */
  train(X: number[][], y: number[]): void {
    if (X.length < 3) return;
    const n = X.length;
    const d = X[0].length;

    // Compute means and stds for standardization
    this.featureMeans = Array(d).fill(0);
    this.featureStds = Array(d).fill(0);

    for (let j = 0; j < d; j++) {
      let sum = 0;
      for (let i = 0; i < n; i++) sum += X[i][j];
      this.featureMeans[j] = sum / n;
    }

    for (let j = 0; j < d; j++) {
      let sum = 0;
      for (let i = 0; i < n; i++) {
        sum += (X[i][j] - this.featureMeans[j]) ** 2;
      }
      this.featureStds[j] = Math.sqrt(sum / n) || 1; // avoid div by zero
    }

    // Standardize features
    const Xs: number[][] = X.map(row =>
      row.map((val, j) => (val - this.featureMeans[j]) / this.featureStds[j]),
    );

    // OLS: w = (X^T X)^-1 X^T y (with bias via augmented matrix)
    // Augment with bias column
    const Xa = Xs.map(row => [...row, 1]);
    const dA = d + 1;

    // X^T X
    const XtX: number[][] = Array.from({ length: dA }, () => Array(dA).fill(0));
    for (let i = 0; i < n; i++) {
      for (let j = 0; j < dA; j++) {
        for (let k = 0; k < dA; k++) {
          XtX[j][k] += Xa[i][j] * Xa[i][k];
        }
      }
    }

    // Add small ridge penalty for stability
    for (let j = 0; j < dA; j++) {
      XtX[j][j] += 1e-6;
    }

    // X^T y
    const Xty: number[] = Array(dA).fill(0);
    for (let i = 0; i < n; i++) {
      for (let j = 0; j < dA; j++) {
        Xty[j] += Xa[i][j] * y[i];
      }
    }

    // Solve via Gauss-Jordan elimination
    const aug = XtX.map((row, i) => [...row, Xty[i]]);
    for (let j = 0; j < dA; j++) {
      // Pivot
      let maxRow = j;
      for (let i = j + 1; i < dA; i++) {
        if (Math.abs(aug[i][j]) > Math.abs(aug[maxRow][j])) maxRow = i;
      }
      [aug[j], aug[maxRow]] = [aug[maxRow], aug[j]];

      const pivot = aug[j][j];
      if (Math.abs(pivot) < 1e-12) continue;

      for (let k = j; k <= dA; k++) aug[j][k] /= pivot;
      for (let i = 0; i < dA; i++) {
        if (i === j) continue;
        const factor = aug[i][j];
        for (let k = j; k <= dA; k++) {
          aug[i][k] -= factor * aug[j][k];
        }
      }
    }

    const solution = aug.map(row => row[dA]);
    this.weights = solution.slice(0, d);
    this.bias = solution[d];
    this.trained = true;
  }

  predict(features: number[]): number | null {
    if (!this.trained || !this.weights) return null;
    // Standardize
    const xs = features.map(
      (val, j) => (val - this.featureMeans[j]) / this.featureStds[j],
    );
    let pred = this.bias;
    for (let j = 0; j < xs.length; j++) {
      pred += this.weights[j] * xs[j];
    }
    return Math.max(0, pred);
  }

  get isTrained(): boolean {
    return this.trained;
  }
}

// ---------------------------------------------------------------------------
// ML Predictor class
// ---------------------------------------------------------------------------

const MIN_SAMPLES = 3;

class MLPredictor {
  private model = new SimpleLinearRegression();
  private data: TrainingSample[] = [];
  private loaded = false;

  async loadData(): Promise<void> {
    if (this.loaded) return;
    try {
      const mlData = await db.getMLData();
      this.data = (mlData || []) as TrainingSample[];
      if (this.data.length >= MIN_SAMPLES) {
        this.trainModel();
      }
      this.loaded = true;
    } catch (e) {
      console.warn('Failed to load ML data:', e);
      this.data = [];
      this.loaded = true;
    }
  }

  private async saveData(): Promise<void> {
    try {
      await db.saveMLData(this.data);
    } catch (e) {
      console.warn('Failed to save ML data:', e);
    }
  }

  private trainModel(): void {
    if (this.data.length < MIN_SAMPLES) return;
    const X = this.data.map(s => [s.sqft, s.bays, s.width, s.height]);
    const y = this.data.map(s => s.cost);
    this.model = new SimpleLinearRegression();
    this.model.train(X, y);
  }

  async addSample(sample: Omit<TrainingSample, 'id' | 'timestamp'>): Promise<boolean> {
    const id = `${sample.project}_${sample.elevation}_${sample.width}_${sample.height}_${sample.bays_wide}x${sample.bays_tall}_${sample.finish}`;

    // Reject duplicates (matches Python behavior)
    if (this.data.some(s => s.id === id)) {
      return false;
    }

    const entry: TrainingSample = {
      ...sample,
      id,
      timestamp: new Date().toISOString(),
    };

    this.data.push(entry);

    if (this.data.length >= MIN_SAMPLES) {
      this.trainModel();
    }
    await this.saveData();
    return true;
  }

  async removeSample(id: string): Promise<void> {
    this.data = this.data.filter(s => s.id !== id);
    if (this.data.length >= MIN_SAMPLES) {
      this.trainModel();
    } else {
      this.model = new SimpleLinearRegression();
    }
    await this.saveData();
  }

  async removeSamplesByProject(project: string): Promise<void> {
    this.data = this.data.filter(s => s.project !== project);
    if (this.data.length >= MIN_SAMPLES) {
      this.trainModel();
    } else {
      this.model = new SimpleLinearRegression();
    }
    await this.saveData();
  }

  async clearAllSamples(): Promise<void> {
    this.data = [];
    this.model = new SimpleLinearRegression();
    await this.saveData();
  }

  /**
   * Check if an elevation is in training by building the full sample ID
   * (matches Python's is_in_training which checks exact ID match).
   */
  isInTraining(
    project: string,
    elevation: string,
    width?: number,
    height?: number,
    baysWide?: number,
    baysTall?: number,
    finish?: string,
  ): boolean {
    if (width != null && height != null && baysWide != null && baysTall != null && finish != null) {
      const id = `${project}_${elevation}_${width}_${height}_${baysWide}x${baysTall}_${finish}`;
      return this.data.some(s => s.id === id);
    }
    // Fallback: match by project + elevation name
    return this.data.some(s => s.project === project && s.elevation === elevation);
  }

  getSampleId(
    project: string,
    elevation: string,
    width?: number,
    height?: number,
    baysWide?: number,
    baysTall?: number,
    finish?: string,
  ): string | null {
    if (width != null && height != null && baysWide != null && baysTall != null && finish != null) {
      const id = `${project}_${elevation}_${width}_${height}_${baysWide}x${baysTall}_${finish}`;
      const found = this.data.find(s => s.id === id);
      return found?.id ?? null;
    }
    const found = this.data.find(s => s.project === project && s.elevation === elevation);
    return found?.id ?? null;
  }

  predict(
    width: number,
    height: number,
    baysWide: number,
    baysTall: number,
    sqft: number,
  ): PredictionResult {
    const bays = baysWide * baysTall || 1;

    // Tier 1: Exact match (within tolerance, exact bays like Python)
    for (const sample of this.data) {
      if (
        Math.abs(sample.width - width) < 1 &&
        Math.abs(sample.height - height) < 1 &&
        sample.bays === bays
      ) {
        return {
          cost: sample.cost,
          confidence: 0.95,
          method: 'exact_match',
          training_samples: this.data.length,
        };
      }
    }

    // Tier 2: ML model prediction
    if (this.model.isTrained) {
      const features = [sqft, bays, width, height];
      const predicted = this.model.predict(features);
      if (predicted != null && predicted > 0) {
        // Confidence formula matches Python: max(0.4, min(0.85, 1 - diff_pct * 0.5))
        const avgCost =
          this.data.reduce((s, d) => s + d.cost, 0) / this.data.length;
        const diffPct = avgCost > 0 ? Math.abs(predicted - avgCost) / avgCost : 1;
        const confidence = Math.max(0.4, Math.min(0.85, 1 - diffPct * 0.5));
        return {
          cost: Math.max(0, Math.round(predicted * 100) / 100),
          confidence: Math.round(confidence * 100) / 100,
          method: 'ml_model',
          training_samples: this.data.length,
        };
      }
    }

    // Tier 3: Average cost per sqft fallback (matches Python: max(1, sqft))
    if (this.data.length > 0) {
      const avgCostPerSqft =
        this.data.reduce((s, d) => s + d.cost / Math.max(1, d.sqft), 0) /
        this.data.length;
      const estimated = sqft > 0 ? sqft * avgCostPerSqft : 5000;
      return {
        cost: Math.round(estimated * 100) / 100,
        confidence: 0.3,
        method: 'avg_per_sqft',
        training_samples: this.data.length,
      };
    }

    // No data fallback
    return {
      cost: Math.max(1000, sqft * 50),
      confidence: 0.1,
      method: 'no_data',
      training_samples: 0,
    };
  }

  async train(): Promise<boolean> {
    if (this.data.length < MIN_SAMPLES) return false;
    this.trainModel();
    return this.model.isTrained;
  }

  getStatistics(): MLStatistics | null {
    if (this.data.length === 0) return null;
    const costs = this.data.map(d => d.cost);
    const avgCost = costs.reduce((s, v) => s + v, 0) / costs.length;
    const avgWidth = this.data.reduce((s, d) => s + d.width, 0) / this.data.length;
    const avgHeight = this.data.reduce((s, d) => s + d.height, 0) / this.data.length;
    const avgSqft = this.data.reduce((s, d) => s + d.sqft, 0) / this.data.length;
    const avgCostPerSqft = avgSqft > 0 ? avgCost / avgSqft : 0;

    // Common configurations
    const configCounts = new Map<string, number>();
    for (const sample of this.data) {
      const config = `${sample.bays_wide}x${sample.bays_tall} bays`;
      configCounts.set(config, (configCounts.get(config) ?? 0) + 1);
    }
    const commonConfigs = Array.from(configCounts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([config, count]) => ({
        config,
        count,
        percentage: (count / this.data.length) * 100,
      }));

    return {
      sample_count: this.data.length,
      avg_cost: avgCost,
      min_cost: Math.min(...costs),
      max_cost: Math.max(...costs),
      avg_width: avgWidth,
      avg_height: avgHeight,
      avg_sqft: avgSqft,
      avg_cost_per_sqft: avgCostPerSqft,
      common_configurations: commonConfigs,
    };
  }

  getStatus(): MLStatus {
    return {
      is_trained: this.model.isTrained,
      sample_count: this.data.length,
      ml_available: true,
    };
  }

  getTrainingData(): TrainingSample[] {
    return [...this.data];
  }
}

// ---------------------------------------------------------------------------
// Singleton
// ---------------------------------------------------------------------------

let _predictor: MLPredictor | null = null;

export function getPredictor(): MLPredictor {
  if (!_predictor) {
    _predictor = new MLPredictor();
  }
  return _predictor;
}
