"""
Simple ML predictor for project cost estimation.
Uses historical project data to predict costs based on dimensions.
"""
import os
import json
from typing import Dict, List, Optional
from datetime import datetime

# Try to import ML libraries
try:
    import numpy as np
    from sklearn.linear_model import LinearRegression
    from sklearn.preprocessing import StandardScaler
    ML_AVAILABLE = True
except ImportError:
    ML_AVAILABLE = False
    print("[ML] scikit-learn or numpy not installed. ML features disabled.")

PROJECTS_DIR = ".files"
ML_DATA_FILE = os.path.join(PROJECTS_DIR, "ml_data.json")


class SimplePredictor:
    """Simple cost predictor based on historical data."""
    
    def __init__(self):
        self.training_data = []
        self.model = None
        self.scaler = None
        self.is_trained = False
        self._load_data()
    
    def _load_data(self, auto_train=True):
        """Load training data from file."""
        if os.path.exists(ML_DATA_FILE):
            try:
                with open(ML_DATA_FILE, 'r') as f:
                    self.training_data = json.load(f)
                print(f"[ML] Loaded {len(self.training_data)} training samples")
                if auto_train and len(self.training_data) >= 3 and ML_AVAILABLE:
                    self._train()
            except Exception as e:
                print(f"[ML] Error loading data: {e}")
                self.training_data = []
    
    def _save_data(self):
        """Save training data to file."""
        try:
            os.makedirs(PROJECTS_DIR, exist_ok=True)
            with open(ML_DATA_FILE, 'w') as f:
                json.dump(self.training_data, f, indent=2)
        except Exception as e:
            print(f"[ML] Error saving data: {e}")
    
    def _train(self):
        """Train the model on current data."""
        if not ML_AVAILABLE or len(self.training_data) < 3:
            self.is_trained = False
            return False
        
        try:
            # Extract features: [sqft, bays, width, height]
            X = []
            y = []
            
            for sample in self.training_data:
                features = [
                    sample.get('sqft', 0),
                    sample.get('bays', 0),
                    sample.get('width', 0),
                    sample.get('height', 0)
                ]
                X.append(features)
                y.append(sample.get('cost', 0))
            
            X = np.array(X)
            y = np.array(y)
            
            # Scale features
            self.scaler = StandardScaler()
            X_scaled = self.scaler.fit_transform(X)
            
            # Train simple linear model
            self.model = LinearRegression()
            self.model.fit(X_scaled, y)
            
            self.is_trained = True
            print(f"[ML] Model trained with {len(self.training_data)} samples")
            return True
            
        except Exception as e:
            print(f"[ML] Training error: {e}")
            self.is_trained = False
            return False
    
    def add_sample(self, elevation_data: Dict, project_name: str = "", elevation_name: str = "") -> bool:
        """Add a training sample from elevation data."""
        # Extract cost from material_impact
        material_impact = elevation_data.get('material_impact', [])
        total_cost = sum(m.get('cost_incurred', 0) for m in material_impact)
        
        if total_cost <= 0:
            return False
        
        # Create unique identifier - include elevation name to distinguish same-dimension elevations
        width = elevation_data.get('opening_width_inches', 0)
        height = elevation_data.get('opening_height_inches', 0)
        bays_w = elevation_data.get('bays_wide', 0)
        bays_h = elevation_data.get('bays_tall', 0)
        sqft = elevation_data.get('total_sqft', 0)
        finish = elevation_data.get('finish', '')
        
        # Include elevation name in ID to allow same-dimension elevations
        sample_id = f"{project_name}_{elevation_name}_{width}_{height}_{bays_w}x{bays_h}_{finish}"
        
        # Check for duplicates
        for existing in self.training_data:
            if existing.get('id') == sample_id:
                return False
        
        sample = {
            'id': sample_id,
            'project': project_name,
            'width': width,
            'height': height,
            'bays': bays_w * bays_h,
            'bays_wide': bays_w,
            'bays_tall': bays_h,
            'sqft': sqft,
            'finish': finish,
            'cost': total_cost,
            'timestamp': datetime.now().isoformat()
        }
        
        self.training_data.append(sample)
        self._save_data()
        
        # Retrain if we have enough data
        if len(self.training_data) >= 3:
            self._train()
        
        return True
    
    def predict(self, elevation_data: Dict) -> Dict:
        """Predict cost for an elevation."""
        width = elevation_data.get('opening_width_inches', 0)
        height = elevation_data.get('opening_height_inches', 0)
        bays_w = elevation_data.get('bays_wide', 0)
        bays_h = elevation_data.get('bays_tall', 0)
        sqft = elevation_data.get('total_sqft', 0)
        
        bays = bays_w * bays_h if bays_w and bays_h else 1
        
        # If we have enough training data, check for exact matches first
        for sample in self.training_data:
            if (abs(sample.get('width', 0) - width) < 1 and 
                abs(sample.get('height', 0) - height) < 1 and
                sample.get('bays', 0) == bays):
                return {
                    'predicted_cost': sample['cost'],
                    'confidence': 0.95,
                    'method': 'exact_match',
                    'training_samples': len(self.training_data)
                }
        
        # Use ML model if trained
        if self.is_trained and self.model and ML_AVAILABLE:
            try:
                features = np.array([[sqft, bays, width, height]])
                features_scaled = self.scaler.transform(features)
                predicted = self.model.predict(features_scaled)[0]
                
                # Calculate confidence based on data variance
                costs = [s['cost'] for s in self.training_data]
                avg_cost = sum(costs) / len(costs)
                
                # Higher confidence if prediction is close to average
                diff_pct = abs(predicted - avg_cost) / avg_cost if avg_cost > 0 else 1
                confidence = max(0.4, min(0.85, 1 - diff_pct * 0.5))
                
                return {
                    'predicted_cost': max(0, round(predicted, 2)),
                    'confidence': round(confidence, 2),
                    'method': 'ml_model',
                    'training_samples': len(self.training_data)
                }
            except Exception as e:
                print(f"[ML] Prediction error: {e}")
        
        # Fallback to simple average-based prediction
        if self.training_data:
            # Find most similar by sqft
            avg_cost_per_sqft = sum(s['cost'] / max(1, s['sqft']) for s in self.training_data) / len(self.training_data)
            estimated = sqft * avg_cost_per_sqft if sqft > 0 else 5000
            
            return {
                'predicted_cost': round(estimated, 2),
                'confidence': 0.3,
                'method': 'avg_per_sqft',
                'training_samples': len(self.training_data)
            }
        
        # No data - rough estimate
        return {
            'predicted_cost': max(1000, sqft * 50),
            'confidence': 0.1,
            'method': 'no_data',
            'training_samples': 0
        }
    
    def get_statistics(self) -> Dict:
        """Get statistics from training data."""
        if not self.training_data:
            return {'error': 'No training data'}
        
        costs = [s['cost'] for s in self.training_data]
        widths = [s['width'] for s in self.training_data]
        heights = [s['height'] for s in self.training_data]
        sqfts = [s['sqft'] for s in self.training_data]
        
        # Count bay configurations
        bay_configs = {}
        for s in self.training_data:
            config = f"{s.get('bays_wide', 0)}x{s.get('bays_tall', 0)}"
            bay_configs[config] = bay_configs.get(config, 0) + 1
        
        # Sort by count
        common_configs = sorted(bay_configs.items(), key=lambda x: x[1], reverse=True)
        
        return {
            'sample_count': len(self.training_data),
            'avg_cost': round(sum(costs) / len(costs), 2),
            'min_cost': round(min(costs), 2),
            'max_cost': round(max(costs), 2),
            'avg_width': round(sum(widths) / len(widths), 0),
            'avg_height': round(sum(heights) / len(heights), 0),
            'avg_sqft': round(sum(sqfts) / len(sqfts), 1),
            'common_configurations': [
                {'config': c[0], 'count': c[1], 'percentage': round(c[1] / len(self.training_data) * 100, 1)}
                for c in common_configs[:5]
            ],
            'is_trained': self.is_trained
        }
    
    def get_status(self) -> Dict:
        """Get ML status."""
        # Reload data to ensure we have the latest count (without auto-training)
        self._load_data(auto_train=False)
        return {
            'training_samples': len(self.training_data),
            'is_trained': self.is_trained,
            'sklearn_available': ML_AVAILABLE
        }
    
    def remove_samples_by_project(self, project_name: str) -> int:
        """Remove all training samples for a given project.
        Returns the number of samples removed.
        """
        initial_count = len(self.training_data)
        self.training_data = [s for s in self.training_data if s.get('project') != project_name]
        removed_count = initial_count - len(self.training_data)
        
        if removed_count > 0:
            # Retrain if we still have enough data, otherwise clear model
            if len(self.training_data) >= 3 and ML_AVAILABLE:
                self._train()
            else:
                self.is_trained = False
                self.model = None
                self.scaler = None
            
            self._save_data()
            print(f"[ML] Removed {removed_count} training samples for project '{project_name}'")
        
        return removed_count
    
    def clear_data(self):
        """Clear all training data."""
        self.training_data = []
        self.is_trained = False
        self.model = None
        self.scaler = None
        self._save_data()


# Global instance
_predictor = None

def get_predictor() -> SimplePredictor:
    """Get or create the predictor instance."""
    global _predictor
    if _predictor is None:
        _predictor = SimplePredictor()
    return _predictor


# Public API
def add_project_to_training(elevation_data: Dict, project_name: str = "", elevation_name: str = "") -> bool:
    """Add an elevation to training data."""
    return get_predictor().add_sample(elevation_data, project_name, elevation_name)

def predict_project_cost(elevation_data: Dict) -> Dict:
    """Predict cost for an elevation."""
    return get_predictor().predict(elevation_data)

def train_ml_model() -> tuple:
    """Train the ML model.
    Returns (success: bool, message: str)
    """
    predictor = get_predictor()
    sample_count = len(predictor.training_data)
    
    if sample_count < 3:
        msg = f"Need minimum of 3 samples to train. Currently have {sample_count} sample(s)."
        print(f"[ML] {msg}")
        return False, msg
    
    success = predictor._train()
    if success:
        return True, f"Model trained successfully with {sample_count} samples!"
    else:
        return False, "Training failed. Please check the logs."

def get_training_status() -> Dict:
    """Get training status."""
    return get_predictor().get_status()

def get_pattern_insights() -> Dict:
    """Get pattern insights from training data."""
    return get_predictor().get_statistics()

def collect_training_data_from_projects() -> int:
    """Collect data from all existing projects."""
    predictor = get_predictor()
    collected = 0
    
    if not os.path.exists(PROJECTS_DIR):
        return 0
    
    for filename in os.listdir(PROJECTS_DIR):
        if filename.endswith('_Elevations.json'):
            project_name = filename.replace('_Elevations.json', '')
            filepath = os.path.join(PROJECTS_DIR, filename)
            
            try:
                with open(filepath, 'r') as f:
                    elevations = json.load(f)
                
                for elev_name, elev_data in elevations.items():
                    if predictor.add_sample(elev_data, project_name):
                        collected += 1
            except Exception as e:
                print(f"[ML] Error processing {filename}: {e}")
    
    print(f"[ML] Collected {collected} new samples")
    return collected

def remove_project_from_training(project_name: str) -> int:
    """Remove all training samples for a project.
    Returns the number of samples removed.
    """
    return get_predictor().remove_samples_by_project(project_name)

def remove_elevation_from_training(elevation_data: Dict, project_name: str = "", elevation_name: str = "") -> bool:
    """Remove a single elevation from training data.
    Returns True if removed, False if not found.
    """
    predictor = get_predictor()
    
    # Create same unique identifier as in add_sample
    width = elevation_data.get('opening_width_inches', 0)
    height = elevation_data.get('opening_height_inches', 0)
    bays_w = elevation_data.get('bays_wide', 0)
    bays_h = elevation_data.get('bays_tall', 0)
    finish = elevation_data.get('finish', '')
    
    sample_id = f"{project_name}_{elevation_name}_{width}_{height}_{bays_w}x{bays_h}_{finish}"
    
    initial_count = len(predictor.training_data)
    predictor.training_data = [s for s in predictor.training_data if s.get('id') != sample_id]
    
    if len(predictor.training_data) < initial_count:
        # Retrain if we still have enough data
        if len(predictor.training_data) >= 3 and ML_AVAILABLE:
            predictor._train()
        else:
            predictor.is_trained = False
            predictor.model = None
            predictor.scaler = None
        predictor._save_data()
        return True
    
    return False

def is_in_training(elevation_data: Dict, project_name: str = "", elevation_name: str = "") -> bool:
    """Check if an elevation is already in training data."""
    predictor = get_predictor()
    
    # Create same unique identifier as in add_sample
    width = elevation_data.get('opening_width_inches', 0)
    height = elevation_data.get('opening_height_inches', 0)
    bays_w = elevation_data.get('bays_wide', 0)
    bays_h = elevation_data.get('bays_tall', 0)
    finish = elevation_data.get('finish', '')
    
    sample_id = f"{project_name}_{elevation_name}_{width}_{height}_{bays_w}x{bays_h}_{finish}"
    
    for sample in predictor.training_data:
        if sample.get('id') == sample_id:
            return True
    
    return False
