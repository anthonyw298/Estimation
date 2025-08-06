from typing import Union

def calculate_rectangle_area(length: float, width: float) -> float:
    """Calculates the area of a rectangle."""
    return length * width

def calculate_perimeter(length: float, width: float) -> float:
    """Calculates the perimeter of a rectangle."""
    return 2 * (length + width)

def convert_inches_to_feet(inches: float) -> float:
    """Converts inches to feet."""
    return inches / 12.0

def convert_feet_to_inches(feet: float) -> float:
    """Converts feet to inches."""
    return feet * 12.0

def calculate_total_gasket_ft(bays_wide: int, bays_tall: int, opening_width: float, opening_height: float, total_count: int) -> float:
    total_inches = (bays_wide * 4 * opening_height) + (bays_tall * 4 * opening_width)
    return (total_inches * total_count) / 12

def calculate_end_dam(total_count: int) -> int:
    return 2 * total_count

def calculate_water_deflector(bays_wide: int, total_count: int) -> int:
    return 2 * bays_wide * total_count

def calculate_assembly_screw(bays_wide: int, bays_tall: int, total_count: int) -> int:
    return ((bays_wide * 8) + ((bays_tall - 1) * 6 * bays_wide)) * total_count

def calculate_sill_flash_screw(bays_wide: int, total_count: int) -> int:
    return 3 * bays_wide * total_count

def calculate_end_dam_screw(total_count: int) -> int:
    return 4 * total_count

def calculate_setting_block_chair(bays_wide: int) -> int:
    return 2 * bays_wide

def calculate_side_block(bays_wide: int, bays_tall: int, total_count: int) -> int:
    return (bays_wide - 1) * bays_tall * total_count

def calculate_setting_block(bays_wide: int, total_count: int) -> int:
    return 2 * bays_wide * total_count

def calculate_anti_walk_block_deep(bays_tall: int, total_count: int) -> int:
    return 2 * bays_tall * total_count

def calculate_anti_walk_block_shallow(bays_wide: int, bays_tall: int, total_count: int) -> int:
    return (bays_wide - 1) * bays_tall * total_count

def calculate_setting_block_int_horizontal(bays_wide: int, total_count: int) -> int:
    return 2 * bays_wide * total_count

def calculate_jamb_ft_v(opening_height: float, total_count: int) -> Union[float, list[float]]:
    """
    Calculates vertical jamb feet. Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2513
    """
    single_instance_qty = (2 * opening_height / 12)
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty

def calculate_sill_ft_h(opening_width: float, total_count: int) -> Union[float, list[float]]:
    """
    Calculates horizontal sill feet. Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2513
    """
    single_instance_qty = (opening_width / 12)
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty

def calculate_flush_filler_v(bays_wide: int, total_count: int, opening_height: float) -> Union[float, list[float]]:
    """
    Calculates vertical flush filler feet. Returns a list if total_count > 1, else a float.
    Associated with profile: E9-2512
    """
    single_instance_qty = ((bays_wide - 1) * opening_height / 12)
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty

def calculate_int_vertical(bays_wide: int, total_count: int, opening_height: float) -> Union[float, list[float]]:
    """
    Calculates intermediate vertical feet. Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2511
    """
    single_instance_qty = ((bays_wide - 1) * opening_height / 12)
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty

def calculate_og_int_horizontal(opening_width: float, total_count: int) -> Union[float, list[float]]:
    """
    Calculates outside glazing intermediate horizontal feet. Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2515
    """
    single_instance_qty = (opening_width / 12)
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty

def calculate_og_head_h(opening_width: float, total_count: int) -> Union[float, list[float]]:
    """
    Calculates outside glazing head horizontal feet. Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2514
    """
    single_instance_qty = (opening_width / 12)
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty

def calculate_sill_flashing_h(opening_width: float, total_count: int) -> Union[float, list[float]]:
    """
    Calculates sill flashing horizontal feet. Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2578
    """
    single_instance_qty = (opening_width / 12)
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty

def calculate_fabrication_joints(bays_wide: int, bays_tall: int, total_count: int) -> int:
    """Calculate number of fabrication joints."""
    return ((4 * bays_wide) + (bays_wide * (2 * (bays_tall - 1))) ) * total_count

def calculate_glass_stop(opening_width: float, bays_tall: int, total_count: int) -> Union[float, list[float]]:
    """
    Calculate glass stop length. Returns a list if total_count > 1, else a float.
    Associated with profile: E9-2519
    """
    single_instance_qty = (opening_width / 12) * bays_tall
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty

def calculate_total_glass(opening_width: float, opening_height: float, total_count: int, bays_wide: int, bays_tall: int) -> float:
    return ((opening_width - (2 * (bays_wide + 1))) * (opening_height - (2 * (bays_tall + 1))) * total_count)/144
def calculate_door_size(door_size_str: str) -> float:
    """
    Calculates the area of a single door from a size string like "3' X 7'".
    Returns the area in square feet.
    """
    try:
        # Normalize the string and split on 'X'
        parts = door_size_str.upper().replace(" ", "").split("X")
        if len(parts) != 2:
            raise ValueError(f"Invalid door size format: {door_size_str}")

        # Remove the apostrophe and convert to float
        width_ft = float(parts[0].replace("'", ""))
        height_ft = float(parts[1].replace("'", ""))

        area = width_ft * height_ft
        return area

    except Exception as e:
        print(f"Error calculating door area for '{door_size_str}': {e}")
        return 0.0
def calculate_door_price(size_str: str, stile: str, hardware_dict: dict) -> float:
    """
    Calculates the total price of a single door unit based on its size, stile, and hardware.
    
    Args:
        size_str (str): Size of the door, e.g., "3' X 8'".
        stile (str): Stile type (e.g., "Narrow", "Medium", "Wide").
        hardware_dict (dict): Dictionary of hardware options, with boolean values.
    
    Returns:
        float: Total price for the door.
    """
    DOOR_BASE_PRICES = {
        '3x7': 1200, '3x8': 1500, '3x9': 1800,
        '6x7': 2400, '6x8': 3000, '6x9': 3600
    }
    STILE_MULTIPLIERS = {'Narrow': 0.9, 'Medium': 1.0, 'Wide': 1.1}
    HARDWARE_PRICE = 69.0

    try:
        parts = size_str.upper().replace("'", "").replace(" ", "").split("X")
        if len(parts) != 2:
            raise ValueError(f"Invalid door size format: {size_str}")

        door_size_key = f"{parts[0]}x{parts[1]}"
        base_price = DOOR_BASE_PRICES.get(door_size_key, 0)

        stile_multiplier = STILE_MULTIPLIERS.get(stile, 1.0)
        door_frame_price = base_price * stile_multiplier

        # Count only selected hardware items (those with value True)
        hardware_count = sum(1 for selected in hardware_dict.values() if selected)

        # Double hardware count for double doors (6' width)
        is_double_door = size_str.strip().startswith("6'")
        if is_double_door:
            hardware_count *= 2

        total_hardware_price = hardware_count * HARDWARE_PRICE

        return door_frame_price + total_hardware_price

    except Exception as e:
        print(f"Error calculating price for door '{size_str}': {e}")
        return 0.0

def calculate_door_info(doors: list) -> list:
    """
    Takes a list of door inputs and returns a list of dictionaries with door information,
    including calculated price and other details.

    Args:
        doors (list): A list of dictionaries, where each dictionary represents a door
                      and contains keys like 'size', 'count', 'stile', and 'hardware'.

    Returns:
        list: A list of dictionaries, where each dictionary represents a door with
              its calculated details and is formatted for the final output.
    """
    door_items = []
    
    if doors:
        for door_info in doors:
            door_size_str = door_info.get('size')
            door_count = door_info.get('count', 0)
            door_stile = door_info.get('stile')
            door_hardware = door_info.get('hardware', [])

            if door_size_str and door_count > 0:
                door_price = calculate_door_price(door_size_str, door_stile, door_hardware)
                
                door_items.append({
                    "description": f"Door ({door_size_str})", 
                    "Style": door_stile,
                    "quantity": door_count,
                    "part_number": "N/A",
                    "type": "Door",
                    'price': door_price,
                    'hardware': door_hardware,
                    'manual': True
                })
    return door_items

def calculate_total_door_area(doors: list) -> float:
    """
    Calculates the total area (in sqft) of all doors in the list.
    Args:
        doors (list): List of door dicts, each with 'size' and 'count'.
    Returns:
        float: Total area in sqft.
    """
    total_area = 0.0
    for door in doors:
        size_str = door.get('size')
        count = door.get('count', 0)
        if size_str and count:
            area = calculate_door_size(size_str)
            total_area += area * count
    return total_area