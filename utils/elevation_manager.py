"""Elevation management operations and data transformation."""
from utils.file_operations import (
    load_elevations,
    save_elevations,
    load_elevation_data
)
from utils.formulas import calculate_rectangle_area, calculate_perimeter
from utils.validation import parse_custom_bays
from systems.yes45tu_front_set import calculate_yes45tu_quantities


def transform_elevation_data_to_ui(elevation_data):
    """Transform elevation data from storage format to UI format.
    
    Args:
        elevation_data: Dictionary with elevation data
    
    Returns:
        dict: Dictionary mapping UI variable keys to their string values
    """
    ui_data = {}
    
    field_mapping = [
        ('system', 'system'),
        ('finish', 'finish'),
        ('total_count', 'total_count'),
        ('bays_wide', 'bays_wide'),
        ('bays_tall', 'bays_tall'),
        ('opening_width_inches', 'opening_width'),
        ('opening_height_inches', 'opening_height'),
        ('custom_bay_widths', 'custom_bay_widths'),
        ('custom_bay_heights', 'custom_bay_heights')
    ]
    
    for storage_key, ui_key in field_mapping:
        if storage_key in elevation_data:
            value = elevation_data[storage_key]
            if storage_key in ['custom_bay_widths', 'custom_bay_heights']:
                # Convert list to comma-separated string
                ui_data[ui_key] = ','.join(map(str, value))
            else:
                ui_data[ui_key] = str(value)
    
    return ui_data


def calculate_elevation_metrics(opening_width, opening_height, total_count):
    """Calculate elevation metrics (sqft, perimeter, etc.).
    
    Args:
        opening_width: Opening width in inches
        opening_height: Opening height in inches
        total_count: Total count of elevations
    
    Returns:
        dict: Dictionary with calculated metrics
    """
    sqft_per = calculate_rectangle_area(opening_width / 12, opening_height / 12)
    total_sqft = sqft_per * total_count
    perimeter = calculate_perimeter(opening_width / 12, opening_height / 12)
    total_perimeter = perimeter * total_count
    
    return {
        'sqft_per_type': sqft_per,
        'total_sqft': total_sqft,
        'perimeter_ft': perimeter,
        'total_perimeter_ft': total_perimeter
    }


def build_elevation_data(
    system, finish, total_count, opening_width, opening_height,
    bays_wide=None, bays_tall=None, custom_bay_widths_str=None,
    custom_bay_heights_str=None, doors=None, system_options=None
):
    """Build elevation data dictionary with all calculations.
    
    Args:
        system: System type
        finish: Finish type
        total_count: Total count
        opening_width: Opening width in inches
        opening_height: Opening height in inches
        bays_wide: Number of bays wide (optional)
        bays_tall: Number of bays tall (optional)
        custom_bay_widths_str: Custom bay widths as comma-separated string (optional)
        custom_bay_heights_str: Custom bay heights as comma-separated string (optional)
        doors: List of door dictionaries (optional)
        system_options: List of system options to check against (optional)
    
    Returns:
        dict: Complete elevation data dictionary
    """
    # Calculate basic metrics
    metrics = calculate_elevation_metrics(opening_width, opening_height, total_count)
    
    elevation_data = {
        'system': system,
        'finish': finish,
        'total_count': total_count,
        'opening_width_inches': opening_width,
        'opening_height_inches': opening_height,
        **metrics
    }
    
    # Calculate system-specific outputs
    calculated_outputs = []
    
    # Check if this is the YES 45TU system
    is_yes45tu = system_options and system == system_options[0] if system_options else False
    if is_yes45tu and bays_wide is not None and bays_tall is not None:
        elevation_data['bays_wide'] = bays_wide
        elevation_data['bays_tall'] = bays_tall
        
        # Parse custom bay dimensions
        custom_bay_widths = parse_custom_bays(
            custom_bay_widths_str or '', opening_width, bays_wide
        )
        custom_bay_heights = parse_custom_bays(
            custom_bay_heights_str or '', opening_height, bays_tall
        )
        
        elevation_data['custom_bay_widths'] = custom_bay_widths
        elevation_data['custom_bay_heights'] = custom_bay_heights
        
        # Calculate system quantities
        calculated_outputs = calculate_yes45tu_quantities(
            bays_wide, bays_tall, total_count,
            opening_width, opening_height, doors or []
        )
    
    elevation_data['calculated_outputs'] = calculated_outputs
    
    return elevation_data


def save_elevation(elevations_json_path, elevation_name, elevation_data):
    """Save an elevation to the elevations JSON file.
    
    Args:
        elevations_json_path: Path to elevations JSON file
        elevation_name: Name of the elevation
        elevation_data: Elevation data dictionary
    """
    elevations = load_elevations(elevations_json_path)
    elevations[elevation_name] = elevation_data
    save_elevations(elevations_json_path, elevations)


def delete_elevation(elevations_json_path, elevation_name):
    """Delete an elevation from the elevations JSON file.
    
    Args:
        elevations_json_path: Path to elevations JSON file
        elevation_name: Name of the elevation to delete
    
    Returns:
        bool: True if elevation was deleted, False if not found
    """
    elevations = load_elevations(elevations_json_path)
    if elevation_name in elevations:
        del elevations[elevation_name]
        save_elevations(elevations_json_path, elevations)
        return True
    return False


def get_elevation_names(elevations_json_path):
    """Get sorted list of elevation names.
    
    Args:
        elevations_json_path: Path to elevations JSON file
    
    Returns:
        list: Sorted list of elevation names
    """
    elevations = load_elevations(elevations_json_path)
    return sorted(elevations.keys())

