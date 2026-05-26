from data.part_number import PART_NUMBER_MAP
from utils.formulas import (
    calculate_end_dam,
    calculate_sill_flash_screw,
    calculate_end_dam_screw,
    calculate_fabrication_joints,
    calculate_total_door_area,
    calculate_glass_to_add_back,
)

from typing import Union


# ---------------------------------------------------------------------------
# Center-set DLO constants (YES 45 TU Center Set / Screw Spline)
# In center set, each vertical contributes the same clearance to every adjacent
# bay (no edge/interior distinction), yielding a uniform 8/3" deduction per bay.
# ---------------------------------------------------------------------------

CS_DLO_DEDUCTION = 8 / 3       # inches — uniform per bay (width and height)
CS_SILL_DEDUCTION = 7 / 16     # inches — additional deduction at bottom row
CS_GLASS_MAKE_ADDITION = 3 / 4  # inches — glass make size = DLO + 3/4"


# ---------------------------------------------------------------------------
# Center-set DLO helpers
# ---------------------------------------------------------------------------

def _cs_dlo_width(bay_width: float) -> float:
    return bay_width - CS_DLO_DEDUCTION


def _cs_dlo_height(bay_height: float, is_bottom: bool) -> float:
    dlo = bay_height - CS_DLO_DEDUCTION
    if is_bottom:
        dlo -= CS_SILL_DEDUCTION
    return dlo


# ---------------------------------------------------------------------------
# Center-set glass & gasket calculations
# ---------------------------------------------------------------------------

def _calculate_cs_total_glass(
    opening_width: float,
    opening_height: float,
    total_count: int,
    bays_wide: int,
    bays_tall: int,
    custom_bay_widths=None,
    custom_bay_heights=None,
) -> float:
    """
    Glass area (sqft) using center-set DLO deductions.
    Uniform 8/3" deduction per bay; extra 7/16" at bottom row.
    Glass make size = DLO + 3/4".
    """
    bay_widths = (
        custom_bay_widths
        if custom_bay_widths and len(custom_bay_widths) == bays_wide
        else [opening_width / bays_wide] * bays_wide
    )
    bay_heights = (
        custom_bay_heights
        if custom_bay_heights and len(custom_bay_heights) == bays_tall
        else [opening_height / bays_tall] * bays_tall
    )

    total_sqft = 0.0
    for col in range(bays_wide):
        dlo_w = _cs_dlo_width(bay_widths[col])
        glass_w = dlo_w + CS_GLASS_MAKE_ADDITION
        for row in range(bays_tall):
            dlo_h = _cs_dlo_height(bay_heights[row], is_bottom=(row == 0))
            glass_h = dlo_h + CS_GLASS_MAKE_ADDITION
            total_sqft += (glass_w * glass_h) / 144

    return total_sqft * total_count


def _calculate_cs_gasket(
    opening_width: float,
    opening_height: float,
    total_count: int,
    bays_wide: int,
    bays_tall: int,
    custom_bay_widths=None,
    custom_bay_heights=None,
) -> float:
    """
    Glazing gasket E2-0052 (ft) = 2 sides × perimeter of each lite.
    Uses glass make sizes (DLO + 3/4") for accurate gasket length.
    """
    bay_widths = (
        custom_bay_widths
        if custom_bay_widths and len(custom_bay_widths) == bays_wide
        else [opening_width / bays_wide] * bays_wide
    )
    bay_heights = (
        custom_bay_heights
        if custom_bay_heights and len(custom_bay_heights) == bays_tall
        else [opening_height / bays_tall] * bays_tall
    )

    total_inches = 0.0
    for col in range(bays_wide):
        dlo_w = _cs_dlo_width(bay_widths[col])
        glass_w = dlo_w + CS_GLASS_MAKE_ADDITION
        for row in range(bays_tall):
            dlo_h = _cs_dlo_height(bay_heights[row], is_bottom=(row == 0))
            glass_h = dlo_h + CS_GLASS_MAKE_ADDITION
            # 2 sides × perimeter of this lite
            total_inches += 2 * 2 * (glass_w + glass_h)

    return (total_inches * total_count) / 12


# ---------------------------------------------------------------------------
# Center-set profile helpers
# ---------------------------------------------------------------------------

def _calculate_cs_vertical_be9_2553(
    opening_height: float, total_count: int, bays_wide: int
) -> list:
    """
    BE9-2553 vertical pieces: 2 jambs + (bays_wide - 1) intermediates = bays_wide + 1 total.
    Each piece is cut to opening_height - 7/16" (sill deduction) — verticals sit on top of sill flashing.
    """
    single_piece_ft = (opening_height - CS_SILL_DEDUCTION) / 12
    pieces_per_elev = bays_wide + 1
    return [single_piece_ft] * (pieces_per_elev * total_count)


def _calculate_cs_head_be9_2553(
    opening_width: float,
    total_count: int,
    bays_wide: int,
    custom_bay_widths: list = None,
) -> list:
    """
    BE9-2553 horizontal head pieces: bays_wide pieces, each cut to DLO width
    (bay_width - 8/3") — horizontals fit between vertical profiles.
    """
    if custom_bay_widths and len(custom_bay_widths) == bays_wide:
        bay_widths_ft = [(w - CS_DLO_DEDUCTION) / 12.0 for w in custom_bay_widths]
    else:
        bay_width_ft = (opening_width / bays_wide - CS_DLO_DEDUCTION) / 12
        bay_widths_ft = [bay_width_ft] * bays_wide

    return bay_widths_ft * total_count


def _calculate_cs_int_horizontal(
    opening_width: float,
    total_count: int,
    bays_wide: int,
    bays_tall: int,
    custom_bay_widths: list = None,
) -> list:
    """
    Horizontal BE9-2556: (bays_tall - 1) * bays_wide pieces, each cut to DLO width
    (bay_width - 8/3") — horizontals fit between vertical profiles.
    """
    if custom_bay_widths and len(custom_bay_widths) == bays_wide:
        bay_widths_ft = [(w - CS_DLO_DEDUCTION) / 12.0 for w in custom_bay_widths]
    else:
        bay_width_ft = (opening_width / bays_wide - CS_DLO_DEDUCTION) / 12
        bay_widths_ft = [bay_width_ft] * bays_wide

    num_rows = bays_tall - 1
    return (bay_widths_ft * num_rows) * total_count


# ---------------------------------------------------------------------------
# Center-set sill / flashing / glass-stop overrides
# These override the shared formulas because center-set cut lengths differ:
#   • Verticals/fillers: height − 7/16" (sill flashing takes up that space)
#   • Sill flashing: width + 1/4" (extends past jambs)
#   • Sill / glass-stop horizontals: DLO width = bay_width − 8/3"
# ---------------------------------------------------------------------------

def _calculate_cs_flush_filler_v(
    bays_wide: int, total_count: int, opening_height: float
) -> list:
    """
    BE9-2552 flush filler verticals: (bays_wide - 1) pieces, each cut to
    opening_height - 7/16" (same deduction as BE9-2553 verticals).
    """
    single_piece_ft = (opening_height - CS_SILL_DEDUCTION) / 12
    pieces_per_elev = bays_wide - 1
    return [single_piece_ft] * (pieces_per_elev * total_count)


def _calculate_cs_sill_ft_h(
    opening_width: float,
    total_count: int,
    bays_wide: int,
    custom_bay_widths: list = None,
) -> list:
    """
    BE9-2579 sill pieces: bays_wide pieces, each cut to DLO width
    (bay_width - 8/3") — same as head/intermediate horizontals.
    """
    if custom_bay_widths and len(custom_bay_widths) == bays_wide:
        bay_widths_ft = [(w - CS_DLO_DEDUCTION) / 12.0 for w in custom_bay_widths]
    else:
        bay_width_ft = (opening_width / bays_wide - CS_DLO_DEDUCTION) / 12
        bay_widths_ft = [bay_width_ft] * bays_wide
    return bay_widths_ft * total_count


def _calculate_cs_sill_flashing_h(opening_width: float, total_count: int) -> list:
    """
    BE9-2578 sill flashing: one piece per elevation at opening_width + 1/4"
    (extends slightly past the jamb profiles).
    """
    single_piece_ft = (opening_width + 0.25) / 12
    return [single_piece_ft] * total_count


def _calculate_cs_glass_stop(
    opening_width: float,
    bays_tall: int,
    total_count: int,
    bays_wide: int,
    custom_bay_widths: list = None,
) -> list:
    """
    E9-1015 glass stop: bays_wide * bays_tall pieces, each cut to DLO width
    (bay_width - 8/3").
    """
    if custom_bay_widths and len(custom_bay_widths) == bays_wide:
        bay_widths_ft = [(w - CS_DLO_DEDUCTION) / 12.0 for w in custom_bay_widths]
    else:
        bay_width_ft = (opening_width / bays_wide - CS_DLO_DEDUCTION) / 12
        bay_widths_ft = [bay_width_ft] * bays_wide
    pieces_per_elev = bay_widths_ft * bays_tall
    return pieces_per_elev * total_count


# ---------------------------------------------------------------------------
# Center-set accessory helpers
# ---------------------------------------------------------------------------

def _calculate_cs_water_deflector(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Water deflector E2-0047: 2 per intermediate horizontal.
    Intermediate horizontals = (bays_tall - 1) * bays_wide.
    """
    return 2 * (bays_tall - 1) * bays_wide * total_count


def _calculate_cs_screw_pc1220(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Assembly screw PC-1220: used at both ends of every horizontal member.
    Members = head (bays_wide) + int horizontals ((bays_tall-1)*bays_wide) + sill (bays_wide)
            = bays_wide * (bays_tall + 1)
    Each member: 2 ends × 2 screws = 4 screws per member.
    Formula: 4 * bays_wide * (bays_tall + 1) * total_count
    """
    return 4 * bays_wide * (bays_tall + 1) * total_count


def _calculate_cs_anti_walk_block(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    W Side Block E2-0153: 2 per lite.
    Formula: 2 * bays_wide * bays_tall * total_count
    """
    return 2 * bays_wide * bays_tall * total_count


def _calculate_cs_setting_block(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Setting / Side Block E2-0020: 2 per lite.
    Formula: 2 * bays_wide * bays_tall * total_count
    """
    return 2 * bays_wide * bays_tall * total_count


def _calculate_cs_flat_filler(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """
    Flat filler E1-1054 at perimeter anchor locations:
      - Sill:  3 * bays_wide  (3 screws/clips per sill bay)
      - Head:  2 * (bays_wide + 1)  (2 per anchor point at each vertical)
      - Jambs: 2 * (bays_tall + 1)  (2 jambs × (bays_tall+1) anchors each)
    Total: 5 * bays_wide + 2 * bays_tall + 4
    """
    return (5 * bays_wide + 2 * bays_tall + 4) * total_count


# ---------------------------------------------------------------------------
# Main calculation function
# ---------------------------------------------------------------------------

def calculate_yes45tu_center_set_quantities(
    bays_wide: int,
    bays_tall: int,
    total_count: int,
    opening_width: float,
    opening_height: float,
    doors=None,
    custom_bay_widths=None,
    custom_bay_heights=None,
    glass_per_sqft=None,
    fabrication_cost_per_joint=None,
) -> list:
    """
    Calculates all output quantities for the 'YES 45TU Center Set' system.
    Returns a list of dicts with description, quantity, part_number, and type.

    Key differences from Front Set (OG):
    - BE9-2553 is used for ALL verticals (jambs + intermediates) AND head pieces
    - Head = bays_wide pieces at bay width (not one full-width piece)
    - No shear blocks (E1-1058/1059) or their screws (PC-1028/FC-1212/PC-1210)
    - No PC-1216 short spline screw
    - No E2-0611 inside setting block
    - DLO uses uniform 8/3" deduction per bay (no edge/interior distinction)
    """
    if doors is None:
        doors = []

    outputs = [
        # --- Accessories ---
        ("E1-0199", calculate_end_dam(total_count)),
        ("E2-0047", _calculate_cs_water_deflector(bays_wide, bays_tall, total_count)),
        ("PC-1220", _calculate_cs_screw_pc1220(bays_wide, bays_tall, total_count)),
        ("PM-1008-SS", calculate_sill_flash_screw(bays_wide, total_count)),
        ("UA-1212", calculate_end_dam_screw(total_count)),
        ("E2-0020", _calculate_cs_setting_block(bays_wide, bays_tall, total_count)),
        ("E2-0153", _calculate_cs_anti_walk_block(bays_wide, bays_tall, total_count)),
        ("E1-1054", _calculate_cs_flat_filler(bays_wide, bays_tall, total_count)),
        # --- Profiles ---
        # BE9-2553 vertical: jambs + intermediate verticals (bays_wide + 1 pieces)
        ("BE9-2553", _calculate_cs_vertical_be9_2553(opening_height, total_count, bays_wide)),
        # BE9-2552: shallow pocket filler verticals (bays_wide - 1 pieces)
        ("BE9-2552", _calculate_cs_flush_filler_v(bays_wide, total_count, opening_height)),
        # BE9-2553 head: bays_wide pieces at bay width
        ("BE9-2553", _calculate_cs_head_be9_2553(opening_width, total_count, bays_wide, custom_bay_widths)),
        # BE9-2579: sill pieces at DLO width
        ("BE9-2579", _calculate_cs_sill_ft_h(opening_width, total_count, bays_wide, custom_bay_widths)),
        # BE9-2556: intermediate horizontals
        ("BE9-2556", _calculate_cs_int_horizontal(opening_width, total_count, bays_wide, bays_tall, custom_bay_widths)),
        # BE9-2578: sill flashing at opening_width + 1/4"
        ("BE9-2578", _calculate_cs_sill_flashing_h(opening_width, total_count)),
        # E9-1015: glass stop at DLO width per lite
        ("E9-1015", _calculate_cs_glass_stop(opening_width, bays_tall, total_count, bays_wide, custom_bay_widths)),
        # E2-0052: glazing gasket
        ("E2-0052", _calculate_cs_gasket(opening_width, opening_height, total_count, bays_wide, bays_tall, custom_bay_widths, custom_bay_heights)),
    ]

    # --- Glass area (center-set DLO formula) ---
    total_glass_area = _calculate_cs_total_glass(
        opening_width, opening_height, total_count, bays_wide, bays_tall,
        custom_bay_widths, custom_bay_heights,
    )

    has_doors = doors and len(doors) > 0
    if has_doors:
        doors_with_total_count = []
        for door in doors:
            door_copy = door.copy()
            door_copy["count"] = door.get("count", 0) * total_count
            doors_with_total_count.append(door_copy)
        total_door_area = calculate_total_door_area(doors_with_total_count)
        total_glass_to_add_back = calculate_glass_to_add_back(doors_with_total_count)
        adjusted_glass_area = max(
            total_glass_area - total_door_area + total_glass_to_add_back, 0
        )
    else:
        total_door_area = 0.0
        total_glass_to_add_back = 0.0
        adjusted_glass_area = total_glass_area

    results = []

    # BE9-2553 appears twice (vertical + head); track counter for labeling
    be9_2553_counter = 0

    for part_number, quantity in outputs:
        desc = None
        part_type = None

        if part_number == "BE9-2553":
            be9_2553_counter += 1
            base_desc = PART_NUMBER_MAP.get("profiles", {}).get("BE9-2553", "UNKNOWN")
            if be9_2553_counter == 1:
                desc = f"Vertical {base_desc}"
            else:
                desc = f"Horizontal {base_desc} (Head)"
            part_type = "profiles"
        else:
            for category, parts_dict in PART_NUMBER_MAP.items():
                if part_number in parts_dict:
                    base_desc = parts_dict[part_number]
                    if part_number in ["BE9-2556", "E9-1015"]:
                        desc = f"Horizontal {base_desc}"
                    elif part_number == "BE9-2552":
                        desc = f"Vertical {base_desc}"
                    elif part_number == "BE9-2579":
                        desc = f"Horizontal {base_desc} (Sill)"
                    else:
                        desc = base_desc
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
            "type": part_type,
        })

    # --- Glass output ---
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

    if has_doors:
        manual_outputs.insert(1, {
            "description": "Door Area (to subtract from glass)",
            "quantity": total_door_area,
            "part_number": "N/A",
            "type": "Calculations",
            "unit": "sqft",
            "manual": True,
        })

    results.extend(manual_outputs)
    return results
