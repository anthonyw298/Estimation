# utils/formulas.py

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

def calculate_jamb_ft_v(opening_height: float, total_count: int) -> float:
    return (2 * opening_height / 12) * total_count

def calculate_sill_ft_h(opening_width: float, total_count: int) -> float:
    return (opening_width / 12) * total_count

def calculate_flush_filler_v(bays_wide: int, total_count: int, opening_height: float) -> float:
    return ((bays_wide-1) * opening_height / 12) * total_count

def calculate_int_vertical(bays_wide: int, total_count: int, opening_height: float) -> float:
    return ((bays_wide-1) * opening_height / 12) * total_count

def calculate_og_int_horizontal(opening_width: float, total_count: int) -> float:
    return (opening_width / 12) * total_count

def calculate_og_head_h(opening_width: float, total_count: int) -> float:
    return (opening_width / 12) * total_count

def calculate_sill_flashing_h(opening_width: float, total_count: int) -> float:
    return (opening_width / 12) * total_count

def calculate_fabrication_joints(bays_wide: int, bays_tall: int, total_count: int) -> int:
    """Calculate number of fabrication joints."""
    return ((4 * bays_wide) + (2 * (bays_tall - 1)) ) * total_count
