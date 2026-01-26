"""
Database abstraction layer - Supabase cloud only (no local JSON).
"""
import os
from typing import Dict, List, Any

# Import supabase
from supabase import create_client, Client

# Import config
from config import SUPABASE_URL, SUPABASE_KEY, LOCAL_DATA_DIR


class Database:
    """Cloud-only database interface for projects, elevations, settings, and doors."""
    
    def __init__(self):
        self.client: Client = None
        self._init_connection()
    
    def _init_connection(self):
        """Initialize Supabase connection."""
        try:
            self.client = create_client(SUPABASE_URL, SUPABASE_KEY)
            print("[DB] Connected to Supabase cloud database")
        except Exception as e:
            print(f"[DB] ERROR: Could not connect to Supabase: {e}")
            raise
        
        # Ensure local directory exists for Excel reports only
        os.makedirs(LOCAL_DATA_DIR, exist_ok=True)
    
    # ==================== PROJECTS ====================
    
    def get_projects(self) -> List[str]:
        """Get list of all project names."""
        result = self.client.table("projects").select("name").execute()
        return [p["name"] for p in result.data]
    
    def save_projects(self, projects: List[str]):
        """Save project list."""
        existing = self.get_projects()
        # Add new projects
        for name in projects:
            if name not in existing:
                self.client.table("projects").insert({"name": name}).execute()
        # Remove deleted projects
        for name in existing:
            if name not in projects:
                self.client.table("projects").delete().eq("name", name).execute()
    
    def delete_project(self, name: str):
        """Delete a project and all its data."""
        self.client.table("projects").delete().eq("name", name).execute()
        self.client.table("elevations").delete().eq("project_name", name).execute()
        self.client.table("settings").delete().eq("project_name", name).execute()
        self.client.table("doors").delete().eq("project_name", name).execute()
        self.client.table("materials").delete().eq("project_name", name).execute()
        
        # Clean up local Excel report if exists
        excel_path = self.get_excel_path(name)
        if os.path.exists(excel_path):
            try: os.remove(excel_path)
            except: pass
    
    # ==================== ELEVATIONS ====================
    
    def get_elevations(self, project_name: str) -> Dict[str, Any]:
        """Get all elevations for a project."""
        result = self.client.table("elevations").select("*").eq("project_name", project_name).execute()
        elevations = {}
        for row in result.data:
            elevations[row["name"]] = row["data"]
        return elevations
    
    def save_elevations(self, project_name: str, elevations: Dict[str, Any]):
        """Save all elevations for a project."""
        for name, data in elevations.items():
            self.client.table("elevations").upsert({
                "project_name": project_name,
                "name": name,
                "data": data
            }, on_conflict="project_name,name").execute()
    
    def delete_elevation(self, project_name: str, elevation_name: str):
        """Delete a specific elevation."""
        self.client.table("elevations").delete().eq("project_name", project_name).eq("name", elevation_name).execute()
        self.client.table("doors").delete().eq("project_name", project_name).eq("elevation_name", elevation_name).execute()
    
    # ==================== SETTINGS ====================
    
    def get_settings(self, project_name: str) -> Dict[str, Any]:
        """Get project settings."""
        result = self.client.table("settings").select("data").eq("project_name", project_name).execute()
        if result.data:
            return result.data[0]["data"]
        return {}
    
    def save_settings(self, project_name: str, settings: Dict[str, Any]):
        """Save project settings."""
        self.client.table("settings").upsert({
            "project_name": project_name,
            "data": settings
        }, on_conflict="project_name").execute()
    
    # ==================== DOORS ====================
    
    def get_doors(self, project_name: str, elevation_name: str) -> List[Dict]:
        """Get doors for an elevation."""
        result = self.client.table("doors").select("data").eq("project_name", project_name).eq("elevation_name", elevation_name).execute()
        if result.data:
            return result.data[0]["data"]
        return []
    
    def save_doors(self, project_name: str, elevation_name: str, doors: List[Dict]):
        """Save doors for an elevation."""
        self.client.table("doors").upsert({
            "project_name": project_name,
            "elevation_name": elevation_name,
            "data": doors
        }, on_conflict="project_name,elevation_name").execute()
    
    # ==================== MATERIALS ====================
    
    def get_materials(self, project_name: str) -> Dict[str, Any]:
        """Get extra materials for a project."""
        result = self.client.table("materials").select("data").eq("project_name", project_name).execute()
        if result.data:
            return result.data[0]["data"]
        return {}
    
    def save_materials(self, project_name: str, materials: Dict[str, Any]):
        """Save extra materials for a project."""
        self.client.table("materials").upsert({
            "project_name": project_name,
            "data": materials
        }, on_conflict="project_name").execute()
    
    # ==================== LOCAL FILES (Excel only) ====================
    
    def get_excel_path(self, project_name: str) -> str:
        """Get path for Excel report (always local - can't store in DB)."""
        clean = project_name.replace(" ", "_").replace("/", "_")
        return os.path.join(LOCAL_DATA_DIR, f"{clean}_Report.xlsx")


# Global database instance
db = Database()
