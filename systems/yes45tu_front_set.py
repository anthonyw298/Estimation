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
    calculate_total_door_area,
    calculate_glass_to_add_back
)

def calculate_yes45tu_quantities(
    bays_wide: int,
    bays_tall: int,
    total_count: int,
    opening_width: float,
    opening_height: float,
    doors=None
) -> list:
    """
    Calculates all the specific output quantities for the 'YES 45TU Front Set(OG)' system
    by calling dedicated formula functions.
    Returns a list of dictionaries with description, quantity, part number, and type.
    """
    # Safety check for doors
    if doors is None:
        doors = []

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
        ("E2-0052", calculate_total_gasket_ft(bays_wide, bays_tall, opening_width, opening_height, total_count))
    ]

    # --- Total area calculations ---
    total_glass_area = calculate_total_glass(opening_width, opening_height, total_count, bays_wide, bays_tall)
    total_door_area = calculate_total_door_area(doors)
    total_glass_to_add_back = calculate_glass_to_add_back(doors)
    adjusted_glass_area = max(total_glass_area - total_door_area + total_glass_to_add_back, 0)  # Prevent negative glass area

    results = []

    # --- Standard outputs ---
    # Explicitly handle the two BE9-2513 entries to label them correctly
    be9_2513_counter = 0

    for part_number, quantity in outputs:
        desc = None
        part_type = None

        if part_number == "BE9-2513":
            be9_2513_counter += 1
            # Fetch base description
            base_desc = PART_NUMBER_MAP.get("profiles", {}).get(part_number, "UNKNOWN")
            if be9_2513_counter == 1:
                desc = f"{base_desc} (Jamb)"
            else:
                desc = f"{base_desc} (Sill)"
            part_type = "profiles"
        else:
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
    
    # Check if the adjusted glass area is zero and add a specific message.
    if adjusted_glass_area == 0:
        glass_output = {
            "description": "Glass Area (Adjusted)",
            "quantity": 0,
            "part_number": "N/A",
            "type": "Glass",
            'price': 0.0,
            'unit': 'sqft',
            'manual': True,
            'message': "Total door area equals or exceeds total glass area. No glass is needed."
        }
    else:
        glass_output = {
            "description": "Glass Area (Adjusted)",
            "quantity": adjusted_glass_area,
            "part_number": "N/A",
            "type": "Glass",
            'price': 10.5,
            'unit': 'sqft',
            'manual': True
        }

    # --- Manual outputs including glass and door area ---
    manual_outputs = [
        glass_output,
        {
            "description": "Door Area (to subtract from glass)",
            "quantity": total_door_area,
            "part_number": "N/A",
            "type": "Calculations",
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
        }
    ]

    results.extend(manual_outputs)
    return results