"""Business validation logic."""
from utils.formulas import calculate_total_door_area, calculate_glass_to_add_back


def parse_custom_bays(input_str, total_dimension, num_bays):
    """Parse custom bay dimensions and distribute remaining dimension equally.
    
    Args:
        input_str: Comma-separated string of custom dimensions
        total_dimension: Total dimension to distribute
        num_bays: Number of bays to create
    
    Returns:
        list: List of bay dimensions
    
    Raises:
        ValueError: If input is invalid
    """
    if not input_str.strip():
        return [total_dimension / num_bays] * num_bays if num_bays > 0 else []
    
    try:
        custom_dims = [float(x) for x in input_str.split(',') if x.strip()]
        if not custom_dims:
            return [total_dimension / num_bays] * num_bays if num_bays > 0 else []
        
        if len(custom_dims) > num_bays:
            raise ValueError(f"Too many dimensions provided. Expected {num_bays}, got {len(custom_dims)}.")
        
        total_custom = sum(custom_dims)
        if total_custom > total_dimension:
            raise ValueError(f"Custom dimensions ({total_custom} in) exceed total dimension ({total_dimension} in).")
        
        remaining_bays = num_bays - len(custom_dims)
        if remaining_bays > 0:
            remaining_dim = (total_dimension - total_custom) / remaining_bays
            return custom_dims + [remaining_dim] * remaining_bays
        return custom_dims
    except ValueError as e:
        raise ValueError(f"Invalid custom bay input: {e}")


def validate_door_addition(glass_area, existing_doors, new_door):
    """Validate that adding a door doesn't reduce glass area to zero or below.
    
    Args:
        glass_area: Total glass area in square feet
        existing_doors: List of existing door dictionaries
        new_door: New door dictionary to add
    
    Returns:
        tuple: (is_valid: bool, error_message: str or None, leftover_glass: float)
    """
    existing_door_area = calculate_total_door_area(existing_doors)
    existing_glass_back = calculate_glass_to_add_back(existing_doors)
    
    new_door_area = calculate_total_door_area([new_door])
    new_glass_back = calculate_glass_to_add_back([new_door])
    
    leftover_glass = glass_area - (existing_door_area + new_door_area) + (existing_glass_back + new_glass_back)
    
    if leftover_glass <= 0:
        return False, "Adding this door reduces glass back area to zero or below.", leftover_glass
    
    return True, None, leftover_glass


def validate_door_update(glass_area, existing_doors, door_index, updated_door):
    """Validate that updating a door doesn't reduce glass area to zero or below.
    
    Args:
        glass_area: Total glass area in square feet
        existing_doors: List of existing door dictionaries
        door_index: Index of the door being updated
        updated_door: Updated door dictionary
    
    Returns:
        tuple: (is_valid: bool, error_message: str or None, leftover_glass: float)
    """
    # Calculate area excluding the door being updated
    doors_excluding_current = existing_doors[:door_index] + existing_doors[door_index+1:]
    
    existing_door_area = calculate_total_door_area(doors_excluding_current)
    existing_glass_back = calculate_glass_to_add_back(doors_excluding_current)
    
    updated_door_area = calculate_total_door_area([updated_door])
    updated_glass_back = calculate_glass_to_add_back([updated_door])
    
    leftover_glass = glass_area - (existing_door_area + updated_door_area) + (existing_glass_back + updated_glass_back)
    
    if leftover_glass <= 0:
        return False, "Updating this door reduces glass back area to zero or below.", leftover_glass
    
    return True, None, leftover_glass


def validate_door_count(door_count_str):
    """Validate door count input.
    
    Args:
        door_count_str: String representation of door count
    
    Returns:
        tuple: (is_valid: bool, error_message: str or None, door_count: int or None)
    """
    if not door_count_str:
        return False, "Door count is required.", None
    
    try:
        door_count = int(door_count_str)
        if door_count <= 0:
            return False, "Door count must be a positive number.", None
        return True, None, door_count
    except ValueError:
        return False, "Door count must be an integer.", None

