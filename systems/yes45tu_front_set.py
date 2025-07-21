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
    calculate_fabrication_joints
)

def calculate_yes45tu_quantities(
    bays_wide: int,
    bays_tall: int,
    total_count: int,
    opening_width: float,
    opening_height: float
) -> list:
    """
    Calculates all the specific output quantities for the 'YES 45TU Front Set(OG)' system
    by calling dedicated formula functions.
    Returns a list of dictionaries with description, quantity, and part number.
    """

    outputs = [
        ("E2-0052", calculate_total_gasket_ft(bays_wide, bays_tall, opening_width, opening_height, total_count)),   # Glazing Gasket For 1” Glazing
        ("E1-0199", calculate_end_dam(total_count)),                                                               # End Dam
        ("E2-0047", calculate_water_deflector(bays_wide, total_count)),                                            # Water Deflector
        ("PC-1220", calculate_assembly_screw(bays_wide, bays_tall, total_count)),                                  # #12 x 1-1/4” PHSMS Assembly Screw
        ("PM-1006-SS", calculate_sill_flash_screw(bays_wide, total_count)),                                        # #10-24 x 3/8” PHMS, Stainless Steel Sill Flash Screw
        ("UA-1212", calculate_end_dam_screw(total_count)),                                                        # #12 x 3/4” UFHSMS End Dam Screw
        ("E1-2530", calculate_setting_block_chair(bays_wide)),                                                    # Setting Block Chair
        ("E2-0166", calculate_side_block(bays_wide, bays_tall, total_count)),                                      # Side Block (Shallow Pocket)
        ("E2-0177", calculate_setting_block(bays_wide, total_count)),                                             # Setting Block (Sill)
        ("E2-0545", calculate_anti_walk_block_deep(bays_tall, total_count)),                                       # 1-1/8” “W” Side Block
        ("E2-0154", calculate_anti_walk_block_shallow(bays_wide, bays_tall, total_count)),                         # 1/2” “W” Side Block
        ("E2-0611", calculate_setting_block_int_horizontal(bays_wide, total_count)),                               # Setting Block (Int. Horizontal)
        ("BE9-2513", calculate_jamb_ft_v(opening_height, total_count)),                                            # Sill/Jamb Screw Spline Assembly (vertical)
        ("BE9-2513", calculate_sill_ft_h(opening_width, total_count)),                                             # Sill/Jamb Screw Spline Assembly (horizontal)
        ("E9-2512", calculate_flush_filler_v(bays_wide, total_count, opening_height)),                             # Flush Filler (Custom Filler)
        ("BE9-2511", calculate_int_vertical(bays_wide, total_count, opening_height)),                              # Two Piece Mullion Screw Spline Assembly
        ("BE9-2515", calculate_og_int_horizontal(opening_width, total_count)),                                     # Horizontal Screw Spline Assembly
        ("BE9-2514", calculate_og_head_h(opening_width, total_count)),                                            # Head Flush Filler
        ("BE9-2578", calculate_sill_flashing_h(opening_width, total_count)),                                       # Thermal Sill Flashing
    ]

    results = []

    for part_number, quantity in outputs:
        desc = None
        part_type = None

        # Search in each category dictionary for the part_number
        for category, parts_dict in PART_NUMBER_MAP.items():
            if part_number in parts_dict:
                desc = parts_dict[part_number]
                part_type = category
                break

        # If not found in any category
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

    return results
