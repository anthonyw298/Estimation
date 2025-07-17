# systems/yes45tu_front_set.py
from utils.part_number import (PART_NUMBER_MAP)
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
        "Total Gasket (Ft)": calculate_total_gasket_ft(bays_wide, bays_tall, opening_width, opening_height, total_count),
        "End Dam": calculate_end_dam(total_count),
        "Water Deflector": calculate_water_deflector(bays_wide, total_count),
        "Assembly Screw": calculate_assembly_screw(bays_wide, bays_tall, total_count),
        "Sill Flash Screw": calculate_sill_flash_screw(bays_wide, total_count),
        "End Dam Screw": calculate_end_dam_screw(total_count),
        "Setting Block Chair": calculate_setting_block_chair(bays_wide),
        "Side Block": calculate_side_block(bays_wide, bays_tall, total_count),
        "Setting Block": calculate_setting_block(bays_wide, total_count),
        "Anti Walk Block Deep Pocket": calculate_anti_walk_block_deep(bays_tall, total_count),
        "Anti Walk Block Shallow Pocket": calculate_anti_walk_block_shallow(bays_wide, bays_tall, total_count),
        "Setting Block (Int. Horizontal)": calculate_setting_block_int_horizontal(bays_wide, total_count),
        "Jamb Ft (V)": calculate_jamb_ft_v(opening_height, total_count),
        "Sill Ft (H)": calculate_sill_ft_h(opening_width, total_count),
        "Flush Filler (V)": calculate_flush_filler_v(bays_wide, total_count, opening_height),
        "Int. Vertical": calculate_int_vertical(bays_wide, total_count, opening_height),
        "OG Int. Horizontal": calculate_og_int_horizontal(opening_width, total_count),
        "OG Head (H)": calculate_og_head_h(opening_width, total_count),
        "Sill Flashing (H)": calculate_sill_flashing_h(opening_width, total_count),
        "Fabrication Joints": calculate_fabrication_joints(bays_wide, bays_tall, total_count)
    }

    results = []
    for key, value in outputs.items():
        part_number = PART_NUMBER_MAP.get(key, "UNKNOWN")
        results.append({
            "description": key,
            "quantity": value,
            "part_number": part_number
        })

    return results