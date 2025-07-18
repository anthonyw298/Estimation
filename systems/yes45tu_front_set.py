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
    calculate_fabrication_joints)

def calculate_yes45tu_quantities(
    bays_wide: int,
    bays_tall: int,
    total_count: int,
    opening_width: float,
    opening_height: float
) -> list:  # 🔄 Note: Return type is now a list!
    """
    Calculates all the specific output quantities for the 'YES 45TU Front Set(OG)' system
    by calling dedicated formula functions.
    Returns a list of dictionaries with description, quantity, and part number.
    """
    outputs = {
    "Glazing Gasket 1”": calculate_total_gasket_ft(bays_wide, bays_tall, opening_width, opening_height, total_count),
    "End Dam": calculate_end_dam(total_count),
    "Water Deflector": calculate_water_deflector(bays_wide, total_count),
    "#12 x 1-1/4” PHSMS": calculate_assembly_screw(bays_wide, bays_tall, total_count),  # Assembly Screw
    "#10-24 x 3/8” PHMS, Stainless Steel": calculate_sill_flash_screw(bays_wide, total_count),  # Sill Flash Screw
    "#12 x 3/4” UFHSMS": calculate_end_dam_screw(total_count),  # End Dam Screw
    "Setting Block Chair": calculate_setting_block_chair(bays_wide),
    "Side Block (Shallow Pocket)": calculate_side_block(bays_wide, bays_tall, total_count),
    "Setting Block (Sill)": calculate_setting_block(bays_wide, total_count),
    "1-1/8” “W” Side Block": calculate_anti_walk_block_deep(bays_tall, total_count),
    "1/2” “W” Side Block": calculate_anti_walk_block_shallow(bays_wide, bays_tall, total_count),
    "Setting Block (Int. Horizontal)": calculate_setting_block_int_horizontal(bays_wide, total_count),
    "Sill/Jamb Screw Spline Assembly": calculate_jamb_ft_v(opening_height, total_count),
    "Sill/Jamb Screw Spline Assembly": calculate_sill_ft_h(opening_width, total_count),
    "Custom Filler": calculate_flush_filler_v(bays_wide, total_count, opening_height),  # E9-2512 unclear
    "Two Piece Mullion Screw Spline Assembly": calculate_int_vertical(bays_wide, total_count, opening_height),
    "Horizontal Screw Spline Assembly": calculate_og_int_horizontal(opening_width, total_count),
    "Head Flush Filler": calculate_og_head_h(opening_width, total_count),
    "Thermal Sill Flashing": calculate_sill_flashing_h(opening_width, total_count),
    "Fabrication Joints": calculate_fabrication_joints(bays_wide, bays_tall, total_count)
}


    results = []

    for desc, qty in outputs.items():
        part_number = None
        part_type = None

        # Search in each category (outer key) for the description
        for outer_key, inner_dict in PART_NUMBER_MAP.items():
            if desc in inner_dict:
                part_number = inner_dict[desc]
                part_type = outer_key
                break

        # If not found, you can decide what to do (e.g., set UNKNOWN)
        if part_number is None:
            part_number = "UNKNOWN"
            part_type = "UNKNOWN"

        results.append({
            "description": desc,
            "quantity": qty,
            "part_number": part_number,
            "type": part_type
        })

    return results
