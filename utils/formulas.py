# utils/all_my_formulas.py

import math

# --- Generic Geometry Formulas (from geometry_formulas.py) ---
def calculate_rectangle_area(length: float, width: float) -> float:
    """Calculates the area of a rectangle."""
    return length * width

def calculate_perimeter(length: float, width: float) -> float:
    """Calculates the perimeter of a rectangle."""
    return 2 * (length + width)

def convert_inches_to_feet(inches: float) -> float:
    """Converts a value from inches to feet."""
    return inches / 12.0

def convert_feet_to_inches(feet: float) -> float:
    """Converts a value from feet to inches."""
    return feet * 12.0

# --- YES 45TU Specific Formulas (from yes45tu_specific_formulas.py) ---
def calculate_total_gasket_ft(bays_wide: int, bays_tall: int, opening_width: float, opening_height: float, total_count: int) -> float:
    """Calculates Total Gasket (Ft) for YES 45TU system."""
    return (((bays_wide * 4 * opening_height) + (bays_tall * 4 * opening_width)) * total_count) / 12

def calculate_end_dam(total_count: int) -> int:
    """Calculates End Dam quantity."""
    return 2 * total_count

def calculate_water_deflector(bays_wide: int, total_count: int) -> int:
    """Calculates Water Deflector quantity."""
    return 2 * bays_wide * total_count

def calculate_assembly_screw(bays_wide: int, bays_tall: int, total_count: int) -> int:
    """Calculates Assembly Screw quantity."""
    return ((bays_wide * 8) + (((bays_tall - 1) * 6) * bays_wide)) * total_count

def calculate_sill_flash_screw(bays_wide: int, total_count: int) -> int:
    """Calculates Sill Flash Screw quantity."""
    return 3 * bays_wide * total_count

def calculate_end_dam_screw(total_count: int) -> int:
    """Calculates End Dam Screw quantity."""
    return 4 * total_count

def calculate_setting_block_chair(bays_wide: int) -> int:
    """Calculates Setting Block Chair quantity (per elevation, not total_count)."""
    return 2 * bays_wide

def calculate_side_block(bays_wide: int, bays_tall: int, total_count: int) -> int:
    """Calculates Side Block quantity."""
    return (bays_wide - 1) * bays_tall * total_count

def calculate_setting_block(bays_wide: int, total_count: int) -> int:
    """Calculates Setting Block quantity."""
    return 2 * bays_wide * total_count

def calculate_anti_walk_block_deep(bays_tall: int, total_count: int) -> int:
    """Calculates Anti Walk Block Deep Pocket quantity."""
    return 2 * bays_tall * total_count

def calculate_anti_walk_block_shallow(bays_wide: int, bays_tall: int, total_count: int) -> int:
    """Calculates Anti Walk Block Shallow Pocket quantity."""
    return (bays_wide - 1) * bays_tall * total_count

def calculate_setting_block_int_horizontal(bays_wide: int, total_count: int) -> int:
    """Calculates Setting Block (Int. Horizontal) quantity."""
    return 2 * bays_wide * total_count

def calculate_jamb_ft_v(opening_height: float, total_count: int) -> float:
    """Calculates Jamb Ft (V) quantity."""
    return (2 * opening_height) / 12 * total_count

def calculate_sill_ft_h(opening_width: float, total_count: int) -> float:
    """Calculates Sill Ft (H) quantity."""
    return (opening_width / 12) * total_count

def calculate_flush_filler_v(bays_wide: int, total_count: int, opening_height: float) -> float:
    """Calculates Flush Filler (V) quantity."""
    return ((bays_wide - 1) * total_count * opening_height) / 12

def calculate_int_vertical(bays_wide: int, total_count: int, opening_height: float) -> float:
    """Calculates Int. Vertical quantity (same as Flush Filler)."""
    return ((bays_wide - 1) * total_count * opening_height) / 12

def calculate_og_int_horizontal(opening_width: float, total_count: int) -> float:
    """Calculates OG Int. Horizontal quantity."""
    return (opening_width / 12) * total_count

def calculate_og_head_h(opening_width: float, total_count: int) -> float:
    """Calculates OG Head (H) quantity."""
    return (opening_width / 12) * total_count

def calculate_sill_flashing_h(opening_width: float, total_count: int) -> float:
    """Calculates Sill Flashing (H) quantity."""
    return (opening_width / 12) * total_count