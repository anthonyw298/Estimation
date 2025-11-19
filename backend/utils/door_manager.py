"""Door management operations."""
from utils.path_manager import get_door_json_path
from utils.file_operations import (
    load_doors,
    save_doors,
    ensure_door_file_exists
)


def get_door_file_path(project_name, elev_type):
    """Get the door JSON file path for an elevation."""
    return get_door_json_path(project_name, elev_type)


def ensure_door_file(project_name, elev_type):
    """Ensure the door JSON file exists for an elevation."""
    door_path = get_door_file_path(project_name, elev_type)
    if door_path:
        ensure_door_file_exists(door_path)
    return door_path


def load_doors_for_elevation(project_name, elev_type):
    """Load doors for a specific elevation.
    
    Returns:
        list: List of door dictionaries
    """
    door_path = get_door_file_path(project_name, elev_type)
    if not door_path:
        return []
    return load_doors(door_path)


def save_doors_for_elevation(project_name, elev_type, doors):
    """Save doors for a specific elevation.
    
    Args:
        project_name: Name of the project
        elev_type: Elevation type name
        doors: List of door dictionaries
    
    Raises:
        ValueError: If door path cannot be determined
        IOError: If save operation fails
    """
    door_path = get_door_file_path(project_name, elev_type)
    if not door_path:
        raise ValueError("Cannot determine door file path. Project name and elevation type are required.")
    save_doors(door_path, doors)


def format_door_for_display(door, index):
    """Format a door dictionary for display in listbox.
    
    Args:
        door: Door dictionary with 'size', 'stile', 'count', 'hardware' keys
        index: 0-based index of the door
    
    Returns:
        str: Formatted string for display
    """
    door_text = f"Door {index+1}: {door['size']}, {door['stile']} Stile, Count: {door['count']}"
    hardware = [hw for hw, var in door['hardware'].items() if var]
    if hardware:
        door_text += f" - Hardware: {', '.join(hardware)}"
    return door_text

