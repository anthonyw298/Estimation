"""File I/O operations for JSON and Excel files."""
import json
import os
from openpyxl import Workbook
from utils.path_manager import MASTER_PROJECT_LIST_FILE


def load_json_file(file_path, default_value=None):
    """Load data from a JSON file."""
    if not os.path.exists(file_path):
        return default_value if default_value is not None else {}
    
    try:
        with open(file_path, 'r') as f:
            return json.load(f)
    except (IOError, json.JSONDecodeError) as e:
        raise IOError(f"Error loading JSON file {file_path}: {e}")


def save_json_file(file_path, data, indent=4):
    """Save data to a JSON file."""
    try:
        with open(file_path, 'w') as f:
            json.dump(data, f, indent=indent)
    except Exception as e:
        raise IOError(f"Error saving JSON file {file_path}: {e}")


def ensure_json_file_exists(file_path, default_data=None):
    """Ensure a JSON file exists, creating it with default data if it doesn't."""
    if not os.path.exists(file_path):
        if default_data is None:
            default_data = {}
        save_json_file(file_path, default_data)


def load_project_list(projects_list_file=MASTER_PROJECT_LIST_FILE):
    """Load the list of projects from the master project list file."""
    return load_json_file(projects_list_file, default_value=[])


def save_project_list(projects, projects_list_file=MASTER_PROJECT_LIST_FILE):
    """Save the list of projects to the master project list file."""
    save_json_file(projects_list_file, projects)


def load_elevations(elevations_json_path):
    """Load elevations data from JSON file."""
    return load_json_file(elevations_json_path, default_value={})


def save_elevations(elevations_json_path, elevations_data):
    """Save elevations data to JSON file."""
    save_json_file(elevations_json_path, elevations_data)


def load_doors(door_json_path):
    """Load doors data from JSON file."""
    if not door_json_path or not os.path.exists(door_json_path):
        return []
    
    try:
        return load_json_file(door_json_path, default_value=[])
    except IOError:
        return []


def save_doors(door_json_path, doors_data):
    """Save doors data to JSON file."""
    if not door_json_path:
        raise ValueError("Door JSON path is required")
    save_json_file(door_json_path, doors_data)


def ensure_door_file_exists(door_json_path):
    """Ensure the door JSON file exists, creating it with empty list if it doesn't."""
    if door_json_path and not os.path.exists(door_json_path):
        ensure_json_file_exists(door_json_path, default_data=[])


def load_elevation_data(elevations_json_path, elevation_name):
    """Load a specific elevation's data from the elevations JSON file."""
    elevations_data = load_elevations(elevations_json_path)
    return elevations_data.get(elevation_name)


def create_excel_file(excel_path):
    """Create a new Excel file with a default 'Report' worksheet."""
    if os.path.exists(excel_path):
        return
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    wb.save(excel_path)

