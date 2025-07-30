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
    calculate_door_size # Your function, correctly imported
)

def calculate_yes45tu_quantities(
    bays_wide: int,
    bays_tall: int,
    total_count: int,
    opening_width: float,
    opening_height: float,
    door_size: str # Type hint changed to str, as it will receive the string '3' X 7''
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
    
    # --- DEBUG PRINTS for Door calculation ---
    print(f"\n--- DEBUG: Door Calculation in yes45tu_front_set.py ---")
    print(f"Input door_size parameter: {door_size} (Type: {type(door_size)})")
    print(f"Total count: {total_count}")
    print(f"Calculated total_glass_area: {total_glass_area:.2f} sqft")

    # --- CORRECTED DOOR AREA CALCULATION ---
    # Pass the door_size string directly to your calculate_door_size function
    # It already handles parsing from 'W' X H'' format.
    if door_size is None or str(door_size).strip().lower() == 'none':
        door_area = 'None'
        print(f"door_size is None or 'None', setting door_area to 'None'.")
    else:
        # Call your existing calculate_door_size function with the string
        door_area = calculate_door_size(door_size, total_count)
        print(f"Calculated door_area from calculate_door_size('{door_size}', {total_count}): {door_area:.2f} sqft")
        
        # If calculate_door_size returns 0.0 (due to internal error handling),
        # we still want to treat it as 'None' for logic consistency (e.g., if input was bad)
        if door_area == 0.0:
            door_area = 'None'
            print("Note: calculate_door_size returned 0.0, treating as 'None' for inclusion logic.")
    # --- END CORRECTED DOOR AREA CALCULATION ---

    results = []

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

    # Determine if the door size exceeds total glass
    if door_area == 'None':
        glass_area_qty = total_glass_area
        door_area_qty = 'None'
        door_desc = "No door size provided"
        print(f"Door not included: {door_desc}")
    elif door_area >= total_glass_area:
        glass_area_qty = total_glass_area
        door_area_qty = 0 # Explicitly set to 0 if it exceeds total glass
        door_desc = "Door size exceeds total glass area"
        print(f"Door quantity set to 0: {door_desc} (Door Area: {door_area:.2f}, Glass Area: {total_glass_area:.2f})")
    else:
        glass_area_qty = total_glass_area - door_area
        door_area_qty = door_area
        door_desc = "Door"
        print(f"Door included. Quantity: {door_area_qty:.2f} sqft. Remaining Glass Area: {glass_area_qty:.2f} sqft")
    
    print(f"Final door_area_qty for 'Door' item dictionary: {door_area_qty}\n")
    # --- END DEBUG PRINTS ---


    # Add manual outputs with possible override
    manual_outputs = [
        {
            "description": "Glass Area",
            "quantity": glass_area_qty,
            "part_number": "N/A",
            "type": "Glass",
            'price': 10.5,
            'unit': 'sqft',
            'manual':True
        },
        {
            "description": "Joints Fabrication Labor",
            "quantity": calculate_fabrication_joints(bays_wide, bays_tall, total_count),
            "part_number": "N/A",
            "type": "Fabrication",
            'price': 15.0,
            'unit': 'joints',
            'manual':True
        },
        {
            "description": "Gasket",
            "quantity": calculate_total_gasket_ft(bays_wide, bays_tall, opening_width, opening_height, total_count),
            "part_number": "E2-0052",
            "type": "Glazing Gasket",
            'unit': 'ft',
            'manual':True
        }
    ]

    if door_area_qty != 'None': # Only add the door if a valid quantity (not 'None') is calculated
        manual_outputs.insert(1, {
            "description": door_desc,
            "quantity": door_area_qty,
            "part_number": "N/A",
            "type": "Door",
            'price': 10, # Price per sqft for door
            'unit': 'sqft',
            'manual': True # Explicitly mark as manual
        })


    results.extend(manual_outputs)
    print(results)
    return results