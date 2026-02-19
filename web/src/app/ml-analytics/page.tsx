'use client';

import { useState, useEffect, useCallback } from 'react';
import { useRouter } from 'next/navigation';
import { db } from '@/lib/database';
import {
  getPredictor,
  type TrainingSample,
  type PredictionResult,
  type MLStatistics,
  type MLStatus,
} from '@/lib/ml-predictor';
import type { ElevationData } from '@/types';
import {
  ArrowLeft,
  Brain,
  RefreshCw,
  Loader2,
  CheckCircle2,
  XCircle,
  Plus,
  Minus,
  BarChart3,
  TrendingUp,
  Activity,
  Layers,
  AlertTriangle,
} from 'lucide-react';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface ElevationEntry {
  project: string;
  elevation: string;
  data: ElevationData;
  prediction: PredictionResult | null;
  isInTraining: boolean;
  cost: number | null; // actual cost if calculated
  selected: boolean;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function fmtCurrency(value: number): string {
  return '$' + value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getConfidenceColor(confidence: number): string {
  if (confidence >= 0.7) return 'text-emerald-400';
  if (confidence >= 0.5) return 'text-yellow-400';
  return 'text-red-400';
}

function getConfidenceBg(confidence: number): string {
  if (confidence >= 0.7) return 'bg-emerald-500/20 text-emerald-400';
  if (confidence >= 0.5) return 'bg-yellow-500/20 text-yellow-400';
  return 'bg-red-500/20 text-red-400';
}

function getMethodLabel(method: string): string {
  switch (method) {
    case 'exact_match': return 'Exact Match';
    case 'ml_model': return 'ML Model';
    case 'avg_per_sqft': return 'Avg/sqft';
    case 'no_data': return 'No Data';
    default: return method;
  }
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export default function MLAnalyticsPage() {
  const router = useRouter();
  const [loading, setLoading] = useState(true);
  const [entries, setEntries] = useState<ElevationEntry[]>([]);
  const [status, setStatus] = useState<MLStatus>({ is_trained: false, sample_count: 0, ml_available: true });
  const [stats, setStats] = useState<MLStatistics | null>(null);
  const [training, setTraining] = useState(false);
  const [loadingProjects, setLoadingProjects] = useState(false);

  // Initialize predictor and auto-load all projects
  useEffect(() => {
    async function init() {
      const predictor = getPredictor();
      await predictor.loadData();
      setStatus(predictor.getStatus());
      setStats(predictor.getStatistics());
      setLoading(false);
      // Auto-load all projects so data is visible immediately
      loadAllProjects();
    }
    init();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Load all projects and their elevations
  const loadAllProjects = useCallback(async () => {
    setLoadingProjects(true);
    try {
      const predictor = getPredictor();
      await predictor.loadData();

      const projects = await db.getProjects();
      const allEntries: ElevationEntry[] = [];

      for (const project of projects) {
        const elevations = await db.getElevations(project);
        for (const [elevName, elev] of Object.entries(elevations)) {
          const width = elev.opening_width_inches || 0;
          const height = elev.opening_height_inches || 0;
          const baysWide = elev.bays_wide || 1;
          const baysTall = elev.bays_tall || 1;
          const sqft = elev.total_sqft || (width * height) / 144;

          // Get actual cost from material_impacts (matches Python's cost source)
          let actualCost: number | null = null;
          if (elev.material_impacts && elev.material_impacts.length > 0) {
            actualCost = elev.material_impacts.reduce(
              (sum, m) => sum + (m.cost_incurred ?? 0),
              0,
            );
          }

          // Get prediction
          let prediction: PredictionResult | null = null;
          if (width > 0 && height > 0) {
            prediction = predictor.predict(width, height, baysWide, baysTall, sqft);
          }

          const isInTraining = predictor.isInTraining(
            project, elevName, width, height, baysWide, baysTall, elev.finish || 'Clear',
          );

          allEntries.push({
            project,
            elevation: elevName,
            data: elev,
            prediction,
            isInTraining,
            cost: actualCost,
            selected: false,
          });
        }
      }

      setEntries(allEntries);
      setStatus(predictor.getStatus());
      setStats(predictor.getStatistics());
    } catch (error) {
      console.error('Failed to load projects:', error);
    } finally {
      setLoadingProjects(false);
    }
  }, []);

  // Train model
  const handleTrain = useCallback(async () => {
    setTraining(true);
    try {
      const predictor = getPredictor();
      const trained = await predictor.train();
      setStatus(predictor.getStatus());
      setStats(predictor.getStatistics());

      if (!trained) {
        alert(`Need at least 3 training samples. Currently have ${predictor.getStatus().sample_count}.`);
      }

      // Re-run predictions with trained model
      if (entries.length > 0) {
        setEntries(prev =>
          prev.map(entry => {
            const width = entry.data.opening_width_inches || 0;
            const height = entry.data.opening_height_inches || 0;
            if (width > 0 && height > 0) {
              const sqft = entry.data.total_sqft || (width * height) / 144;
              return {
                ...entry,
                prediction: predictor.predict(width, height, entry.data.bays_wide || 1, entry.data.bays_tall || 1, sqft),
              };
            }
            return entry;
          }),
        );
      }
    } finally {
      setTraining(false);
    }
  }, [entries]);

  // Add/remove from training
  const toggleTraining = useCallback(async (entry: ElevationEntry) => {
    const predictor = getPredictor();
    const w = entry.data.opening_width_inches || 0;
    const h = entry.data.opening_height_inches || 0;
    const bw = entry.data.bays_wide || 1;
    const bt = entry.data.bays_tall || 1;
    const fin = entry.data.finish || 'Clear';
    if (entry.isInTraining) {
      const sampleId = predictor.getSampleId(entry.project, entry.elevation, w, h, bw, bt, fin);
      if (sampleId) await predictor.removeSample(sampleId);
    } else {
      if (entry.cost == null || entry.cost === 0) {
        alert('This elevation has no calculated cost. Run "Calculate & Save" first.');
        return;
      }
      const width = entry.data.opening_width_inches || 0;
      const height = entry.data.opening_height_inches || 0;
      const sqft = entry.data.total_sqft || (width * height) / 144;
      await predictor.addSample({
        project: entry.project,
        elevation: entry.elevation,
        width,
        height,
        bays: (entry.data.bays_wide || 1) * (entry.data.bays_tall || 1),
        bays_wide: entry.data.bays_wide || 1,
        bays_tall: entry.data.bays_tall || 1,
        sqft,
        finish: entry.data.finish || 'Clear',
        cost: entry.cost,
      });
    }

    setEntries(prev =>
      prev.map(e =>
        e.project === entry.project && e.elevation === entry.elevation
          ? { ...e, isInTraining: !e.isInTraining }
          : e,
      ),
    );
    setStatus(predictor.getStatus());
    setStats(predictor.getStatistics());
  }, []);

  // Bulk operations
  const selectAll = () => setEntries(prev => prev.map(e => ({ ...e, selected: true })));
  const deselectAll = () => setEntries(prev => prev.map(e => ({ ...e, selected: false })));

  const addSelectedToTraining = useCallback(async () => {
    const predictor = getPredictor();
    const selected = entries.filter(e => e.selected && !e.isInTraining && e.cost != null && e.cost > 0);

    for (const entry of selected) {
      const width = entry.data.opening_width_inches || 0;
      const height = entry.data.opening_height_inches || 0;
      const sqft = entry.data.total_sqft || (width * height) / 144;
      await predictor.addSample({
        project: entry.project,
        elevation: entry.elevation,
        width,
        height,
        bays: (entry.data.bays_wide || 1) * (entry.data.bays_tall || 1),
        bays_wide: entry.data.bays_wide || 1,
        bays_tall: entry.data.bays_tall || 1,
        sqft,
        finish: entry.data.finish || 'Clear',
        cost: entry.cost!,
      });
    }

    setEntries(prev =>
      prev.map(e => {
        if (e.selected && !e.isInTraining && e.cost != null && e.cost > 0) {
          return { ...e, isInTraining: true, selected: false };
        }
        return { ...e, selected: false };
      }),
    );
    setStatus(predictor.getStatus());
    setStats(predictor.getStatistics());
  }, [entries]);

  const removeSelectedFromTraining = useCallback(async () => {
    const predictor = getPredictor();
    const selected = entries.filter(e => e.selected && e.isInTraining);

    for (const entry of selected) {
      const w = entry.data.opening_width_inches || 0;
      const h = entry.data.opening_height_inches || 0;
      const bw = entry.data.bays_wide || 1;
      const bt = entry.data.bays_tall || 1;
      const fin = entry.data.finish || 'Clear';
      const sampleId = predictor.getSampleId(entry.project, entry.elevation, w, h, bw, bt, fin);
      if (sampleId) await predictor.removeSample(sampleId);
    }

    setEntries(prev =>
      prev.map(e => {
        if (e.selected && e.isInTraining) {
          return { ...e, isInTraining: false, selected: false };
        }
        return { ...e, selected: false };
      }),
    );
    setStatus(predictor.getStatus());
    setStats(predictor.getStatistics());
  }, [entries]);

  const addAllToTraining = useCallback(async () => {
    const predictor = getPredictor();
    const eligible = entries.filter(e => !e.isInTraining && e.cost != null && e.cost > 0);

    for (const entry of eligible) {
      const width = entry.data.opening_width_inches || 0;
      const height = entry.data.opening_height_inches || 0;
      const sqft = entry.data.total_sqft || (width * height) / 144;
      await predictor.addSample({
        project: entry.project,
        elevation: entry.elevation,
        width,
        height,
        bays: (entry.data.bays_wide || 1) * (entry.data.bays_tall || 1),
        bays_wide: entry.data.bays_wide || 1,
        bays_tall: entry.data.bays_tall || 1,
        sqft,
        finish: entry.data.finish || 'Clear',
        cost: entry.cost!,
      });
    }

    setEntries(prev =>
      prev.map(e => {
        if (!e.isInTraining && e.cost != null && e.cost > 0) {
          return { ...e, isInTraining: true, selected: false };
        }
        return { ...e, selected: false };
      }),
    );
    setStatus(predictor.getStatus());
    setStats(predictor.getStatistics());
  }, [entries]);

  // Loading
  if (loading) {
    return (
      <div className="min-h-screen bg-[#06060a] flex items-center justify-center">
        <div className="flex flex-col items-center">
          <Loader2 className="w-8 h-8 text-[#3b82f6] animate-spin mb-4" />
          <p className="text-[#8b8d9a] text-sm">Loading ML Analytics...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#06060a] flex flex-col">
      {/* Header */}
      <header className="glass border-b border-[#1e1e2a]/60 bg-[#06060a]/80 backdrop-blur-sm sticky top-0 z-30 flex-shrink-0">
        <div className="px-6 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <button
              onClick={() => router.push('/')}
              className="p-2 rounded-lg hover:bg-[#111118] text-[#8b8d9a] hover:text-[#eeeef2] transition-colors"
            >
              <ArrowLeft className="w-5 h-5" />
            </button>
            <div className="w-px h-6 bg-[#1e1e2a]" />
            <div className="flex items-center gap-2">
              <Brain className="w-5 h-5 text-purple-500" />
              <h1 className="text-base font-semibold text-[#eeeef2] tracking-tight">ML Analytics</h1>
            </div>
          </div>

          <div className="flex items-center gap-3">
            {/* Status */}
            <div className="flex items-center gap-2 text-xs text-[#55566a]">
              <span className="transition-all duration-200">
                Model: {status.is_trained ? (
                  <span className="text-emerald-400">Trained</span>
                ) : (
                  <span className="text-[#55566a]">Not trained</span>
                )}
              </span>
              <span className="text-[#2a2a3a]">|</span>
              <span>Samples: {status.sample_count}</span>
            </div>

            <button
              onClick={loadAllProjects}
              disabled={loadingProjects}
              className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-[#8b8d9a] hover:text-[#eeeef2] hover:bg-[#111118] rounded-lg transition-all duration-200 disabled:opacity-40"
            >
              {loadingProjects ? <Loader2 className="w-4 h-4 animate-spin" /> : <RefreshCw className="w-4 h-4" />}
              Load Projects
            </button>

            <button
              onClick={handleTrain}
              disabled={training || status.sample_count < 3}
              className="flex items-center gap-2 px-4 py-2 text-sm font-medium bg-purple-600 hover:bg-purple-500 active:scale-[0.97] text-white rounded-lg transition-all duration-200 shadow-md shadow-purple-500/10 disabled:opacity-40"
            >
              {training ? <Loader2 className="w-4 h-4 animate-spin" /> : <Brain className="w-4 h-4" />}
              Train Model
            </button>
          </div>
        </div>

        {/* Warning banner */}
        {status.sample_count > 0 && status.sample_count < 3 && (
          <div className="px-6 py-2 bg-amber-900/15 border-t border-amber-500/15">
            <div className="flex items-center gap-2 text-xs text-yellow-400">
              <AlertTriangle className="w-4 h-4" />
              Need at least 3 training samples to train the model. Currently have {status.sample_count}.
            </div>
          </div>
        )}
      </header>

      {/* Body - Two Column Layout */}
      <div className="flex flex-1 overflow-hidden">
        {/* Left Panel - Project Predictions */}
        <div className="flex-1 flex flex-col overflow-hidden border-r border-[#1e1e2a]">
          <div className="px-4 py-3 border-b border-[#1e1e2a] bg-[#0a0a10]">
            <div className="flex items-center justify-between">
              <h2 className="text-xs font-semibold text-[#55566a] uppercase tracking-wider">
                Project Predictions ({entries.length})
              </h2>
              {entries.length > 0 && (
                <div className="flex items-center gap-2">
                  <button onClick={selectAll} className="text-xs text-[#55566a] hover:text-[#c4c5cf] transition-colors duration-200">Select All</button>
                  <span className="text-[#2a2a3a]">|</span>
                  <button onClick={deselectAll} className="text-xs text-[#55566a] hover:text-[#c4c5cf] transition-colors duration-200">Deselect</button>
                  <span className="text-[#2a2a3a]">|</span>
                  <button
                    onClick={addSelectedToTraining}
                    className="text-xs text-emerald-500 hover:text-emerald-400 transition-colors duration-200"
                  >
                    Add Selected
                  </button>
                  <span className="text-[#2a2a3a]">|</span>
                  <button
                    onClick={addAllToTraining}
                    className="text-xs text-blue-500 hover:text-blue-400 transition-colors duration-200"
                  >
                    Add All
                  </button>
                  <span className="text-[#2a2a3a]">|</span>
                  <button
                    onClick={removeSelectedFromTraining}
                    className="text-xs text-red-500 hover:text-red-400 transition-colors duration-200"
                  >
                    Remove Selected
                  </button>
                </div>
              )}
            </div>
          </div>

          <div className="flex-1 overflow-y-auto py-1">
            {entries.length === 0 ? (
              <div className="flex flex-col items-center justify-center h-full text-center px-6 animate-fade-up opacity-0" style={{ animationFillMode: 'forwards', animationDelay: '0.1s' }}>
                <Layers className="w-10 h-10 text-[#3e3f4d] mb-4" />
                <p className="text-sm text-[#8b8d9a] mb-2">No projects loaded</p>
                <p className="text-xs text-[#3e3f4d] max-w-xs">
                  Click &ldquo;Load Projects&rdquo; to fetch all projects and their elevations for prediction analysis.
                </p>
              </div>
            ) : (
              entries.map((entry, idx) => (
                <div
                  key={`${entry.project}-${entry.elevation}`}
                  className={`flex items-center gap-3 px-4 py-3 mx-2 my-1 rounded-xl border transition-all duration-200 ${
                    entry.selected
                      ? 'border-purple-500/30 bg-purple-500/5'
                      : 'border-[#1e1e2a] hover:bg-[#111118] hover:border-[#2a2a3a]'
                  }`}
                >
                  {/* Checkbox */}
                  <input
                    type="checkbox"
                    checked={entry.selected}
                    onChange={() =>
                      setEntries(prev =>
                        prev.map((e, i) => (i === idx ? { ...e, selected: !e.selected } : e)),
                      )
                    }
                    className="h-4 w-4 rounded border-[#2a2a3a] bg-[#0c0c12] text-purple-500 accent-purple-500 flex-shrink-0"
                  />

                  {/* Info */}
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-medium text-[#eeeef2] truncate">
                        {entry.project} / {entry.elevation}
                      </span>
                      {entry.isInTraining && (
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400 flex-shrink-0" />
                      )}
                    </div>
                    <div className="text-xs text-[#55566a] mt-0.5">
                      {entry.data.opening_width_inches}&Prime; x {entry.data.opening_height_inches}&Prime;
                      {' | '}
                      {entry.data.bays_wide}x{entry.data.bays_tall} bays
                      {' | '}
                      {entry.data.finish}
                    </div>
                  </div>

                  {/* Prediction / Cost */}
                  <div className="text-right flex-shrink-0">
                    {entry.cost != null ? (
                      <div>
                        <p className="text-sm font-mono font-bold text-[#eeeef2] tabular-nums">
                          {fmtCurrency(entry.cost)}
                        </p>
                        <p className="text-[10px] text-[#55566a]">actual</p>
                      </div>
                    ) : entry.prediction ? (
                      <div>
                        <p className={`text-sm font-mono font-bold tabular-nums ${getConfidenceColor(entry.prediction.confidence)}`}>
                          {fmtCurrency(entry.prediction.cost)}
                        </p>
                        <p className={`text-[10px] ${getConfidenceColor(entry.prediction.confidence)}`}>
                          {(entry.prediction.confidence * 100).toFixed(0)}% - {getMethodLabel(entry.prediction.method)}
                        </p>
                      </div>
                    ) : (
                      <p className="text-xs text-[#3e3f4d]">No data</p>
                    )}
                  </div>

                  {/* Add/remove button */}
                  <button
                    onClick={() => toggleTraining(entry)}
                    className={`p-1.5 rounded-md transition-colors duration-200 flex-shrink-0 ${
                      entry.isInTraining
                        ? 'text-red-400 hover:bg-red-500/10'
                        : 'text-emerald-400 hover:bg-emerald-500/10'
                    }`}
                    title={entry.isInTraining ? 'Remove from training' : 'Add to training'}
                  >
                    {entry.isInTraining ? <Minus className="w-4 h-4" /> : <Plus className="w-4 h-4" />}
                  </button>
                </div>
              ))
            )}
          </div>
        </div>

        {/* Right Panel - Pattern Insights */}
        <div className="w-[380px] flex-shrink-0 bg-[#0a0a10] overflow-y-auto p-5 space-y-5">
          <h2 className="text-xs font-semibold text-[#55566a] uppercase tracking-wider">
            Pattern Insights
          </h2>

          {!stats || !status.is_trained ? (
            <div className="text-center py-16">
              <Brain className="w-10 h-10 text-[#3e3f4d] mx-auto mb-4" />
              <p className="text-sm text-[#8b8d9a] mb-1">
                {!stats ? 'No training data yet' : 'Model not trained'}
              </p>
              <p className="text-xs text-[#3e3f4d]">
                {!stats
                  ? 'Add elevations to training, then train the model to see pattern insights.'
                  : `${stats.sample_count} sample(s) loaded. Click "Train Model" to see pattern insights.`}
              </p>
            </div>
          ) : (
            <>
              {/* Model Status */}
              <div className={`rounded-xl border px-4 py-3 transition-all duration-200 ${
                status.is_trained
                  ? 'border-emerald-500/30 bg-emerald-500/5'
                  : 'border-yellow-500/30 bg-yellow-500/5'
              }`}>
                <div className="flex items-center gap-2">
                  {status.is_trained ? (
                    <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                  ) : (
                    <XCircle className="w-4 h-4 text-yellow-400" />
                  )}
                  <span className={`text-sm font-medium ${status.is_trained ? 'text-emerald-400' : 'text-yellow-400'}`}>
                    {status.is_trained ? 'Model Trained' : 'Model Not Trained'}
                  </span>
                  <span className="text-xs text-[#55566a] ml-auto">{stats.sample_count} samples</span>
                </div>
              </div>

              {/* Statistics */}
              <div className="bg-[#111118] border border-[#1e1e2a] rounded-xl p-4 space-y-3 shadow-lg shadow-black/20">
                <div className="flex items-center gap-2 mb-2">
                  <BarChart3 className="w-4 h-4 text-blue-500" />
                  <h3 className="text-sm font-semibold text-[#eeeef2]">Statistics</h3>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div className="bg-[#08080e] border border-[#1e1e2a] rounded-lg p-2.5">
                    <p className="text-[10px] text-[#55566a] uppercase">Avg Cost</p>
                    <p className="text-sm font-mono font-bold text-[#eeeef2] tabular-nums">
                      {fmtCurrency(stats.avg_cost)}
                    </p>
                  </div>
                  <div className="bg-[#08080e] border border-[#1e1e2a] rounded-lg p-2.5">
                    <p className="text-[10px] text-[#55566a] uppercase">Avg SQFT</p>
                    <p className="text-sm font-mono font-bold text-[#eeeef2] tabular-nums">
                      {stats.avg_sqft.toFixed(1)}
                    </p>
                  </div>
                  <div className="bg-[#08080e] border border-[#1e1e2a] rounded-lg p-2.5">
                    <p className="text-[10px] text-[#55566a] uppercase">Avg Size</p>
                    <p className="text-sm font-mono text-[#eeeef2] tabular-nums">
                      {stats.avg_width.toFixed(0)}&Prime; x {stats.avg_height.toFixed(0)}&Prime;
                    </p>
                  </div>
                  <div className="bg-[#08080e] border border-[#1e1e2a] rounded-lg p-2.5">
                    <p className="text-[10px] text-[#55566a] uppercase">Avg $/sqft</p>
                    <p className="text-sm font-mono font-bold text-emerald-400 tabular-nums">
                      {fmtCurrency(stats.avg_cost_per_sqft)}
                    </p>
                  </div>
                </div>
              </div>

              {/* Cost Range */}
              <div className="bg-[#111118] border border-[#1e1e2a] rounded-xl p-4 space-y-3 shadow-lg shadow-black/20">
                <div className="flex items-center gap-2 mb-2">
                  <TrendingUp className="w-4 h-4 text-emerald-500" />
                  <h3 className="text-sm font-semibold text-[#eeeef2]">Cost Range</h3>
                </div>

                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-[10px] text-[#55566a] uppercase">Min</p>
                    <p className="text-sm font-mono text-red-400 tabular-nums">
                      {fmtCurrency(stats.min_cost)}
                    </p>
                  </div>
                  <div className="flex-1 mx-4 h-1.5 bg-[#1e1e2a] rounded-full overflow-hidden">
                    <div
                      className="h-full bg-gradient-to-r from-red-500 via-yellow-500 to-emerald-500 rounded-full"
                      style={{ width: '100%' }}
                    />
                  </div>
                  <div className="text-right">
                    <p className="text-[10px] text-[#55566a] uppercase">Max</p>
                    <p className="text-sm font-mono text-emerald-400 tabular-nums">
                      {fmtCurrency(stats.max_cost)}
                    </p>
                  </div>
                </div>
              </div>

              {/* Common Configurations */}
              {stats.common_configurations.length > 0 && (
                <div className="bg-[#111118] border border-[#1e1e2a] rounded-xl p-4 space-y-3 shadow-lg shadow-black/20">
                  <div className="flex items-center gap-2 mb-2">
                    <Activity className="w-4 h-4 text-cyan-500" />
                    <h3 className="text-sm font-semibold text-[#eeeef2]">Common Configurations</h3>
                  </div>

                  <div className="space-y-2">
                    {stats.common_configurations.map((config, i) => (
                      <div key={i} className="flex items-center justify-between">
                        <span className="text-xs text-[#c4c5cf]">{config.config}</span>
                        <div className="flex items-center gap-2">
                          <div className="w-16 h-1.5 bg-[#1e1e2a] rounded-full overflow-hidden">
                            <div
                              className="h-full bg-cyan-500 rounded-full"
                              style={{ width: `${config.percentage}%` }}
                            />
                          </div>
                          <span className="text-xs text-[#55566a] font-mono tabular-nums w-14 text-right">
                            {config.count} ({config.percentage.toFixed(0)}%)
                          </span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
