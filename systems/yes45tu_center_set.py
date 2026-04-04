from data.part_number import PART_NUMBER_MAP
from utils.formulas import (
    calculate_total_gasket_ft,
    calculate_end_dam,
    calculate_sill_flash_screw,
    calculate_end_dam_screw,
    calculate_jamb_ft_v,
    calculate_sill_ft_h,
    calculate_flush_filler_v,
    calculate_int_vertical,
    calculate_sill_flashing_h,
    calculate_glass_stop,
    calculate_total_glass,
    calculate_fabrication_joints,
    calculate_total_door_area,
    calculate_glass_to_add_back,
)

from typing import Union


# ---------------------------------------------------------------------------
# Center-set-specific private helper formulas
# ---------------------------------------------------------------------------


def _calculate_shear_block_sill_horizontal(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Shear block E1-1058 for sill and horizontal members.
    Needed at each end of sill and horizontal members where they attach to verticals.
    Formula: 2 * bays_wide * bays_tall * total_count
    (sill contributes 2 * bays_wide, each horizontal row contributes 2 * bays_wide,
     total horizontal rows = bays_tall - 1, so combined = 2 * bays_wide * bays_tall)
    """
    return 2 * bays_wide * bays_tall * total_count


def _calculate_shear_block_head(bays_wide: int, total_count: int) -> int:
    """
    Shear block E1-1059 for head members.
    Needed at each end of head members.
    Formula: 2 * bays_wide * total_count
    """
    return 2 * bays_wide * total_count


def _calculate_cs_water_deflector(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Water deflector for center set: 2 per intermediate horizontal.
    Intermediate horizontals = (bays_tall - 1) * bays_wide.
    Formula: 2 * (bays_tall - 1) * bays_wide * total_count
    """
    return 2 * (bays_tall - 1) * bays_wide * total_count


def _calculate_cs_screw_pc1216(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Assembly screw PC-1216 (short spline) for center set.
    Used where horizontals/sills/head connect to INTERMEDIATE verticals.
    Each intermediate vertical has (bays_tall + 1) connection points
    (1 head + (bays_tall-1) horizontals + 1 sill), each with 2 screws on each side.
    Formula: 4 * (bays_wide - 1) * (bays_tall + 1) * total_count
    """
    return 4 * (bays_wide - 1) * (bays_tall + 1) * total_count


def _calculate_cs_screw_pc1220(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Assembly screw PC-1220 (long spline) for center set.
    Used at every end of head/horizontal/sill members connecting to verticals.
    Total members = 1 head + (bays_tall-1)*bays_wide horizontals + bays_wide sills
                   = 1 + bays_wide * bays_tall
    Each member has 2 ends, each with 2 screws.
    Formula: 4 * (1 + bays_wide * bays_tall) * total_count
    """
    return 4 * (1 + bays_wide * bays_tall) * total_count


def _calculate_cs_anti_walk_block(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Anti-walk block E2-0153 for center set: 1 per lite (in deep vertical pocket).
    Formula: bays_wide * bays_tall * total_count
    """
    return bays_wide * bays_tall * total_count


def _calculate_cs_setting_block_outside(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Setting block outside E2-0020: 2 per lite at DLO quarter points.
    Formula: 2 * bays_wide * bays_tall * total_count
    """
    return 2 * bays_wide * bays_tall * total_count


def _calculate_cs_setting_block_inside(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Setting block inside E2-0611: 2 per lite (center set has both inside and outside).
    Formula: 2 * bays_wide * bays_tall * total_count
    """
    return 2 * bays_wide * bays_tall * total_count


def _calculate_cs_flat_filler(bays_wide: int, total_count: int) -> int:
    """
    Flat filler E1-1054: used at head and jamb anchor locations.
    2 per bay at head + 2 per jamb (2 jambs).
    Formula: (2 * bays_wide + 4) * total_count
    """
    return (2 * bays_wide + 4) * total_count


def _calculate_cs_head_h(
    opening_width: float, total_count: int
) -> Union[float, list[float]]:
    """
    Head BE9-2553 for center set: 1 piece at FULL frame width (not per bay).
    Returns a list with a single full-width piece per unit.
    """
    full_width_ft = opening_width / 12
    if total_count > 1:
        return [full_width_ft] * total_count
    return [full_width_ft]


def _calculate_cs_int_horizontal(
    opening_width: float,
    total_count: int,
    bays_wide: int,
    bays_tall: int,
    custom_bay_widths: list = None,
) -> Union[float, list[float]]:
    """
    Horizontal BE9-2556 for center set:
    (bays_tall - 1) * bays_wide pieces, each cut to bay width.
    """
    if bays_wide and bays_wide > 0:
        if custom_bay_widths and len(custom_bay_widths) == bays_wide:
            bay_widths_ft = [w / 12.0 for w in custom_bay_widths]
        else:
            bay_width_ft = (opening_width / bays_wide) / 12
            bay_widths_ft = [bay_width_ft] * bays_wide

        # Repeat bay widths for each intermediate horizontal row
        num_rows = bays_tall - 1
        pieces = bay_widths_ft * num_rows

        if total_count > 1:
            return pieces * total_count
        return pieces

    # Fallback: single piece at full width per row
    single_piece_ft = opening_width / 12
    num_pieces = (bays_tall - 1) * total_count
    return [single_piece_ft] * num_pieces


def calculate_yes45tu_center_set_quantities(
    bays_wide: int,
    bays_tall: int,
    total_count: int,
    opening_width: float,
    opening_height: float,
    doors=None,
    custom_bay_widths=None,
    glass_per_sqft=None,
    fabrication_cost_per_joint=None,
) -> list:
    """
    Calculates all the specific output quantities for the 'YES 45TU Center Set' system
    by calling dedicated formula functions.
    Returns a list of dictionaries with description, quantity, part number, and type.
    """
    # Safety check for doors
    if doors is None:
        doors = []

    # --- Shear block counts (used for screw calculations below) ---
    shear_block_sill_horiz_count = _calculate_shear_block_sill_horizontal(
        bays_wide, bays_tall, total_count
    )
    shear_block_head_count = _calculate_shear_block_head(bays_wide, total_count)

    outputs = [
        # --- Accessories ---
        ("E1-0199", calculate_end_dam(total_count)),
        ("E2-0047", _calculate_cs_water_deflector(bays_wide, bays_tall, total_count)),
        ("PC-1216", _calculate_cs_screw_pc1216(bays_wide, bays_tall, total_count)),
        ("PC-1220", _calculate_cs_screw_pc1220(bays_wide, bays_tall, total_count)),
        ("PM-1008-SS", calculate_sill_flash_screw(bays_wide, total_count)),
        ("UA-1212", calculate_end_dam_screw(total_count)),
        ("E2-0153", _calculate_cs_anti_walk_block(bays_wide, bays_tall, total_count)),
        ("E2-0020", _calculate_cs_setting_block_outside(bays_wide, bays_tall, total_count)),
        ("E2-0611", _calculate_cs_setting_block_inside(bays_wide, bays_tall, total_count)),
        ("E1-1054", _calculate_cs_flat_filler(bays_wide, total_count)),
        # --- Shear blocks ---
        ("E1-1058", shear_block_sill_horiz_count),
        ("E1-1059", shear_block_head_count),
        # --- Shear block screws ---
        ("PC-1028", 2 * shear_block_sill_horiz_count),
        ("FC-1212", shear_block_head_count),
        ("PC-1210", shear_block_sill_horiz_count),
        # --- Profiles (center set part numbers) ---
        ("BE9-2551", calculate_jamb_ft_v(opening_height, total_count)),
        (
            "BE9-2579",
            calculate_sill_ft_h(
                opening_width, total_count, bays_wide, custom_bay_widths
            ),
        ),
        ("BE9-2552", calculate_flush_filler_v(bays_wide, total_count, opening_height)),
        ("BE9-2555", calculate_int_vertical(bays_wide, total_count, opening_height)),
        (
            "BE9-2556",
            _calculate_cs_int_horizontal(
                opening_width, total_count, bays_wide, bays_tall, custom_bay_widths
            ),
        ),
        (
            "BE9-2553",
            _calculate_cs_head_h(opening_width, total_count),
        ),
        ("BE9-2578", calculate_sill_flashing_h(opening_width, total_count)),
        (
            "E9-1015",
            calculate_glass_stop(
                opening_width, bays_tall, total_count, bays_wide, custom_bay_widths
            ),
        ),
        (
            "E2-0052",
            calculate_total_gasket_ft(
                bays_wide, bays_tall, opening_width, opening_height, total_count
            ),
        ),
    ]

    # --- Total area calculations ---
    total_glass_area = calculate_total_glass(
        opening_width,
        opening_height,
        total_count,
        bays_wide,
        bays_tall,
        custom_bay_widths,
    )

    # Only calculate door area if doors exist
    # Door count is per elevation, so multiply by total_count for total calculations
    has_doors = doors and len(doors) > 0
    if has_doors:
        # Multiply door counts by total_count since door count is per elevation
        doors_with_total_count = []
        for door in doors:
            door_copy = door.copy()
            door_copy["count"] = door.get("count", 0) * total_count
            doors_with_total_count.append(door_copy)
        total_door_area = calculate_total_door_area(doors_with_total_count)
        total_glass_to_add_back = calculate_glass_to_add_back(doors_with_total_count)
        adjusted_glass_area = max(
            total_glass_area - total_door_area + total_glass_to_add_back, 0
        )  # Prevent negative glass area
    else:
        total_door_area = 0.0
        total_glass_to_add_back = 0.0
        adjusted_glass_area = total_glass_area  # No door adjustments needed

    results = []

    # --- Standard outputs ---
    for part_number, quantity in outputs:
        desc = None
        part_type = None

        for category, parts_dict in PART_NUMBER_MAP.items():
            if part_number in parts_dict:
                base_desc = parts_dict[part_number]
                # Add Horizontal/Vertical prefix for specific profile parts
                if part_number in ["BE9-2553", "BE9-2556", "E9-1015"]:
                    desc = f"Horizontal {base_desc}"
                elif part_number in ["BE9-2552", "BE9-2555"]:
                    desc = f"Vertical {base_desc}"
                elif part_number == "BE9-2551":
                    desc = f"Vertical {base_desc} (Jamb)"
                elif part_number == "BE9-2579":
                    desc = f"Horizontal {base_desc}"
                elif part_number == "E2-0153":
                    desc = f"Anti-Walk Block {base_desc}"
                else:
                    desc = base_desc
                part_type = category
                break

        if desc is None:
            desc = "UNKNOWN"
            part_type = "UNKNOWN"
            part_number = "UNKNOWN"

        results.append(
            {
                "description": desc,
                "quantity": quantity,
                "part_number": part_number,
                "type": part_type,
            }
        )

    # Check if the adjusted glass area is zero and add a specific message.
    if adjusted_glass_area == 0:
        glass_output = {
            "description": "Glass Area (Adjusted)",
            "quantity": 0,
            "part_number": "N/A",
            "type": "Glass",
            "price": 0.0,
            "unit": "sqft",
            "manual": True,
            "message": "Total door area equals or exceeds total glass area. No glass is needed.",
        }
    else:
        glass_output = {
            "description": "Glass Area (Adjusted)",
            "quantity": adjusted_glass_area,
            "part_number": "N/A",
            "type": "Glass",
            "price": float(glass_per_sqft) if glass_per_sqft is not None else 10.5,
            "unit": "sqft",
            "manual": True,
        }

    fab_price = (
        float(fabrication_cost_per_joint)
        if fabrication_cost_per_joint is not None
        else 15.0
    )
    manual_outputs = [
        glass_output,
        {
            "description": "Joints Fabrication Labor",
            "quantity": calculate_fabrication_joints(bays_wide, bays_tall, total_count),
            "part_number": "N/A",
            "type": "Fabrication",
            "price": fab_price,
            "unit": "joints",
            "manual": True,
        },
    ]

    # Only include door area calculation if doors exist
    if has_doors:
        manual_outputs.insert(
            1,
            {
                "description": "Door Area (to subtract from glass)",
                "quantity": total_door_area,
                "part_number": "N/A",
                "type": "Calculations",
                "unit": "sqft",
                "manual": True,
            },
        )

    results.extend(manual_outputs)
    return results
