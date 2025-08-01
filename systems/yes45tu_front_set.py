# systems/yes45tu_front_set.py
from data.part_number import (PART_NUMBER_MAP)
from utils.formulas import (
    calculate_total_gasket_ft,
    calculate_end_dam,
    calculate_water_deflector,
    calculate_assembly_screw,
    calculate_sill_flash_screw,
    calculate_end_dam_screw,
    calculate_setting_block_chair,
    calculate_side_block,
    calculate_setting_block,
    calculate_anti_walk_block_deep,
    calculate_anti_walk_block_shallow,
    calculate_setting_block_int_horizontal,
    calculate_jamb_ft_v,
    calculate_sill_ft_h,
    calculate_flush_filler_v,
    calculate_int_vertical,
    calculate_og_int_horizontal,
    calculate_og_head_h,
    calculate_sill_flashing_h,
    calculate_glass_stop,
    calculate_total_glass,
    calculate_fabrication_joints,
    calculate_door_size,
    calculate_door_price
)

def calculate_yes45tu_quantities(
    bays_wide: int,
    bays_tall: int,
    total_count: int,
    opening_width: float,
    opening_height: float,
    doors: list
) -> list:
    """
    Calculates all the specific output quantities for the 'YES 45TU Front Set(OG)' system
    by calling dedicated formula functions.
    Returns a list of dictionaries with description, quantity, part number, and type.
    """

    outputs = [
        ("E1-0199", calculate_end_dam(total_count)),
        ("E2-0047", calculate_water_deflector(bays_wide, total_count)),
        ("PC-1220", calculate_assembly_screw(bays_wide, bays_tall, total_count)),
        ("PM-1006-SS", calculate_sill_flash_screw(bays_wide, total_count)),
        ("UA-1212", calculate_end_dam_screw(total_count)),
        ("E1-2530", calculate_setting_block_chair(bays_wide)),
        ("E2-0166", calculate_side_block(bays_wide, bays_tall, total_count)),
        ("E2-0177", calculate_setting_block(bays_wide, total_count)),
        ("E2-0545", calculate_anti_walk_block_deep(bays_tall, total_count)),
        ("E2-0154", calculate_anti_walk_block_shallow(bays_wide, bays_tall, total_count)),
        ("E2-0611", calculate_setting_block_int_horizontal(bays_wide, total_count)),
        ("BE9-2513", calculate_jamb_ft_v(opening_height, total_count)),
        ("BE9-2513", calculate_sill_ft_h(opening_width, total_count)),
        ("E9-2512", calculate_flush_filler_v(bays_wide, total_count, opening_height)),
        ("BE9-2511", calculate_int_vertical(bays_wide, total_count, opening_height)),
        ("BE9-2515", calculate_og_int_horizontal(opening_width, total_count)),
        ("BE9-2514", calculate_og_head_h(opening_width, total_count)),
        ("BE9-2578", calculate_sill_flashing_h(opening_width, total_count)),
        ("E9-2519", calculate_glass_stop(opening_width, bays_tall, total_count)),
    ]

    total_glass_area = calculate_total_glass(opening_width, opening_height, total_count, bays_wide, bays_tall)
    
    # --- Door Calculation ---
    total_door_area = 0
    door_items = []
    
    if doors:
        for door_info in doors:
            door_size_str = door_info.get('size')
            door_count = door_info.get('count', 0)
            door_stile = door_info.get('stile')
            door_hardware = door_info.get('hardware', [])

            if door_size_str and door_count > 0:
                door_area = calculate_door_size(door_size_str)
                
                # Calculate the price for this specific door
                door_price = calculate_door_price(door_size_str, door_stile, door_hardware)
                
                total_door_area += door_area
                
                door_items.append({
                    "description": f"Door ({door_size_str}) - Stile: {door_stile}",
                    "quantity": door_count,
                    "part_number": "N/A",
                    "type": "Door",
                    'price': door_price,
                    'hardware': door_hardware,
                    'manual': True
                })

    # Adjust glass area based on total door area
    glass_area_qty = max(0, total_glass_area - total_door_area)

    results = []

    # First, append the calculated doors
    results.extend(door_items)

    # Then, append the other outputs from the `outputs` list
    for part_number, quantity in outputs:
        desc = None
        part_type = None

        for category, parts_dict in PART_NUMBER_MAP.items():
            if part_number in parts_dict:
                desc = parts_dict[part_number]
                part_type = category
                break

        if desc is None:
            desc = "UNKNOWN"
            part_type = "UNKNOWN"
            part_number = "UNKNOWN"

        results.append({
            "description": desc,
            "quantity": quantity,
            "part_number": part_number,
            "type": part_type
        })

    # Finally, append the remaining manual outputs
    manual_outputs = [
        {
            "description": "Glass Area",
            "quantity": glass_area_qty,
            "part_number": "N/A",
            "type": "Glass",
            'price': 10.5,
            'unit': 'sqft',
            'manual': True
        },
        {
            "description": "Joints Fabrication Labor",
            "quantity": calculate_fabrication_joints(bays_wide, bays_tall, total_count),
            "part_number": "N/A",
            "type": "Fabrication",
            'price': 15.0,
            'unit': 'joints',
            'manual': True
        },
        {
            "description": "Gasket",
            "quantity": calculate_total_gasket_ft(bays_wide, bays_tall, opening_width, opening_height, total_count),
            "part_number": "E2-0052",
            "type": "Glazing Gasket",
            'unit': 'ft',
            'manual': True
        }
    ]

    results.extend(manual_outputs)
    
    print(results, 'this is results')
    return results

# The `calculate_door_size` and `calculate_door_price` functions are not part of the issue, but they are included here for completeness.
def calculate_door_size(door_size_str: str) -> float:
    """
    Calculates the area of a single door from a size string like "3' X 7'".
    Returns the area in square feet.
    """
    try:
        parts = door_size_str.upper().replace(" ", "").split("X")
        if len(parts) != 2:
            raise ValueError(f"Invalid door size format: {door_size_str}")

        width_ft = float(parts[0].replace("'", ""))
        height_ft = float(parts[1].replace("'", ""))

        area = width_ft * height_ft
        return area

    except Exception as e:
        print(f"Error calculating door area for '{door_size_str}': {e}")
        return 0.0

def calculate_door_price(size_str: str, stile: str, hardware_list: list) -> float:
    """
    Calculates the total price of a single door unit based on its size, stile, and hardware.
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
        
        hardware_count = len(hardware_list)
        is_double_door = size_str.startswith("6'")
        if is_double_door:
            hardware_count *= 2
        
        total_hardware_price = hardware_count * HARDWARE_PRICE
        
        return door_frame_price + total_hardware_price

    except Exception as e:
        print(f"Error calculating price for door '{size_str}': {e}")
        return 0.0