from typing import Union
import re

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

def calculate_door_price(size_str: str, width_type: str, hardware_dict: dict, finish: str) -> float:
    """
    Calculates total price of a door given size (e.g. "3' X 8'"), width_type ("Narrow", "Medium", "Wide"),
    finish ("Clear", "Black", "Paint"), and selected hardware.
    Hardware prices also vary by finish.
    """

    # Full door price matrix
    DOOR_PRICES = {
        "3x7": {
            "Narrow": {"Clear": 880.00, "Black": 1035.00, "Paint": 1269.00},
            "Medium": {"Clear": 1180.00, "Black": 1245.00, "Paint": 1653.00},
            "Wide": {"Clear": 1304.25, "Black": 1413.75, "Paint": 1744.50}
        },
        "3x8": {
            "Narrow": {"Clear": 921.75, "Black": 1083.00, "Paint": 1328.25},
            "Medium": {"Clear": 1235.25, "Black": 1304.25, "Paint": 1727.25},
            "Wide": {"Clear": 1365.00, "Black": 1479.00, "Paint": 1825.50}
        },
        "3x9": {
            "Narrow": {"Clear": 986.25, "Black": 1159.50, "Paint": 1422.75},
            "Medium": {"Clear": 1321.50, "Black": 1395.75, "Paint": 1849.50},
            "Wide": {"Clear": 1461.75, "Black": 1584.00, "Paint": 1953.00}
        },
        "6x7": {
            "Narrow": {"Clear": 1715.25, "Black": 1863.75, "Paint": 2657.25},
            "Medium": {"Clear": 2310.75, "Black": 2445.75, "Paint": 3156.75},
            "Wide": {"Clear": 2559.00, "Black": 2781.00, "Paint": 3435.75}
        },
        "6x8": {
            "Narrow": {"Clear": 1812.00, "Black": 1970.25, "Paint": 2799.75},
            "Medium": {"Clear": 2442.00, "Black": 2589.75, "Paint": 3338.25},
            "Wide": {"Clear": 2700.00, "Black": 2932.50, "Paint": 3630.00}
        },
        "6x9": {
            "Narrow": {"Clear": 1943.25, "Black": 2111.25, "Paint": 2988.00},
            "Medium": {"Clear": 2624.25, "Black": 2782.50, "Paint": 3564.00},
            "Wide": {"Clear": 2901.00, "Black": 3150.00, "Paint": 3861.00}
        }
    }

    # Hardware base prices by finish
    HARDWARE_PRICES = {
        "Concealed Closer": {"Clear": 473.00, "Black": 473.00, "Paint": 473.00},
        "Exit Device": {"Clear": 475.00, "Black": 475.00, "Paint": 475.00},  # Will adjust for >7 ft doors below
        "Continuous Hinges": {"Clear": 285.00, "Black": 375.00, "Paint": 375.00},
        "Lever Handle & Latch Lock": {"Clear": 334.00, "Black": 334.00, "Paint": 334.00},
        "Latchlock with Paddle": {"Clear": 433.00, "Black": 433.00, "Paint": 433.00},
        "Electric Strike": {"Clear": 355.00, "Black": 355.00, "Paint": 355.00}
    }

    try:
        # Normalize size key (e.g., "3' X 8'" → "3x8")
        parts = size_str.upper().replace("'", "").replace(" ", "").split("X")
        if len(parts) != 2:
            raise ValueError(f"Invalid door size format: {size_str}")
        door_size_key = f"{parts[0].lower()}x{parts[1].lower()}"

        # Get base price
        base_price = DOOR_PRICES[door_size_key][width_type][finish]

        # Calculate hardware cost based on finish
        hardware_total = 0
        for hw, selected in hardware_dict.items():
            if selected and hw in HARDWARE_PRICES:
                price = HARDWARE_PRICES[hw][finish]

                # Special rule: Exit Device is $550 for all doors except 3x7 and 6x7
                if hw == "Exit Device" and door_size_key not in ["3x7", "6x7"]:
                    price = 550.00

                # Double price for double doors (all hardware)
                if door_size_key.startswith("6x"):
                    price *= 2

                hardware_total += price

        return base_price + hardware_total

    except KeyError as e:
        print(f"Invalid key in pricing lookup: {e}")
        return 0.0
    except Exception as e:
        print(f"Error calculating price: {e}")
        return 0.0

def calculate_door_info(doors: list,finish='Clear') -> list:
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
    print(doors,'Hello this is door list')
    if doors:
        for door_info in doors:
            door_size_str = door_info.get('size')
            door_count = door_info.get('count', 0)
            door_stile = door_info.get('stile')
            door_hardware = door_info.get('hardware', [])

            if door_size_str and door_count > 0:
                door_price = calculate_door_price(door_size_str, door_stile, door_hardware,finish)
                
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

def calculate_glass_to_add_back(doors):
    """
    Calculate total glass back area in sqft based on door sizes.

    Args:
        doors (list of dict): List of doors with 'size', 'count', 'stile' keys.

    Returns:
        float: Total glass back area in sqft.
    """

    deductions = {
        'Narrow': {'height': 13.5625, 'width': 4.875},
        'Medium': {'height': 15.1875, 'width': 7.625},
        'Wide': {'height': 16.25, 'width': 10.625},
    }

    if not doors or not isinstance(doors, list):
        return 0

    total_area = 0.0

    for door in doors:
        stile = door.get('stile', '').title()
        if stile not in deductions:
            continue

        count = door.get('count', 1)
        size_str = door.get('size', '')
        if not size_str:
            continue

        # Parse door opening width and height in feet from 'size' string
        size_match = re.match(r"\s*(\d+)' *[xX] *(\d+)'", size_str)
        if not size_match:
            continue

        opening_width_ft = int(size_match.group(1))
        opening_height_ft = int(size_match.group(2))

        # Convert to inches
        opening_width_in = opening_width_ft * 12
        opening_height_in = opening_height_ft * 12

        # For 6' width doors, width deduction applies after dividing width by 2 (paired door)
        if opening_width_ft == 6:
            glass_width = (opening_width_in / 2) - deductions[stile]['width']
        else:
            glass_width = opening_width_in - deductions[stile]['width']

        glass_height = opening_height_in - deductions[stile]['height']

        if glass_width <= 0 or glass_height <= 0:
            continue

        area_sqft = (glass_width * glass_height) / 144
        total_area += area_sqft * count

    return round(total_area, 2)
