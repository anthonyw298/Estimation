from typing import Union
import re


def calculate_rectangle_area(length: float, width: float) -> float:
    """Calculates the area of a rectangle."""
    return length * width


def calculate_perimeter(length: float, width: float) -> float:
    """Calculates the perimeter of a rectangle."""
    return 2 * (length + width)


def calculate_total_gasket_ft(
    bays_wide: int,
    bays_tall: int,
    opening_width: float,
    opening_height: float,
    total_count: int,
) -> float:
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


def calculate_setting_block_chair(bays_wide: int, total_count: int) -> int:
    return 2 * bays_wide * total_count


def calculate_side_block(bays_wide: int, bays_tall: int, total_count: int) -> int:
    return (bays_wide - 1) * bays_tall * total_count


def calculate_setting_block(bays_wide: int, total_count: int) -> int:
    return 2 * bays_wide * total_count


def calculate_anti_walk_block_deep(bays_tall: int, total_count: int) -> int:
    return 2 * bays_tall * total_count


def calculate_anti_walk_block_shallow(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    return (bays_wide - 1) * bays_tall * total_count


def calculate_setting_block_int_horizontal(bays_wide: int, total_count: int) -> int:
    return 2 * bays_wide * total_count


def calculate_jamb_ft_v(
    opening_height: float, total_count: int
) -> Union[float, list[float]]:
    """
    Calculates vertical jamb feet. Returns a list of full height (split into 2 pieces - left and right).
    Associated with profile: BE9-2513 (jamb uses full height, 2 pieces)
    """
    # Full height: each jamb piece is opening_height (left and right sides)
    single_piece_qty = opening_height / 12  # Convert inches to feet
    # Always return as list with 2 pieces per instance (left and right)
    if total_count > 1:
        return [single_piece_qty, single_piece_qty] * total_count
    return [single_piece_qty, single_piece_qty]


def calculate_sill_ft_h(
    opening_width: float,
    total_count: int,
    bays_wide: int = None,
    custom_bay_widths: list = None,
) -> Union[float, list[float]]:
    """
    Calculates horizontal sill feet. Returns a list of bay widths if bays_wide is provided,
    otherwise returns the whole width. Returns a list if total_count > 1, else a float.
    Uses custom_bay_widths if provided, otherwise divides equally.
    Associated with profile: BE9-2513 (sill only uses bay widths)
    """
    if bays_wide and bays_wide > 0:
        # Use custom bay widths if provided, otherwise divide equally
        if custom_bay_widths and len(custom_bay_widths) == bays_wide:
            # Convert custom bay widths from inches to feet
            bay_widths_ft = [w / 12.0 for w in custom_bay_widths]
        else:
            # Equal division
            bay_width_ft = (opening_width / bays_wide) / 12
            bay_widths_ft = [bay_width_ft] * bays_wide

        if total_count > 1:
            return bay_widths_ft * total_count
        return bay_widths_ft

    single_instance_qty = opening_width / 12
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty


def calculate_flush_filler_v(
    bays_wide: int, total_count: int, opening_height: float
) -> Union[float, list[float]]:
    """
    Calculates vertical flush filler feet. Returns a list of full height for each filler.
    Associated with profile: E9-2512 (uses full height, one piece per filler)
    """
    # Full height: each filler piece is opening_height
    single_piece_qty = opening_height / 12  # Convert inches to feet
    # Number of fillers = (bays_wide - 1), one piece each
    pieces_per_instance = bays_wide - 1
    if total_count > 1:
        return [single_piece_qty] * pieces_per_instance * total_count
    return [single_piece_qty] * pieces_per_instance


def calculate_int_vertical(
    bays_wide: int, total_count: int, opening_height: float
) -> Union[float, list[float]]:
    """
    Calculates intermediate vertical feet. Returns a list of full height for each mullion.
    Associated with profile: BE9-2511 (uses full height, one piece per mullion)
    """
    # Full height: each mullion piece is opening_height
    single_piece_qty = opening_height / 12  # Convert inches to feet
    # Number of mullions = (bays_wide - 1), one piece each
    pieces_per_instance = bays_wide - 1
    if total_count > 1:
        return [single_piece_qty] * pieces_per_instance * total_count
    return [single_piece_qty] * pieces_per_instance


def calculate_og_int_horizontal(
    opening_width: float,
    total_count: int,
    bays_wide: int = None,
    custom_bay_widths: list = None,
) -> Union[float, list[float]]:
    """
    Calculates outside glazing intermediate horizontal feet. Returns a list of bay widths if bays_wide is provided,
    otherwise returns the whole width. Uses custom_bay_widths if provided, otherwise divides equally.
    Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2515 (uses bay widths)
    """
    if bays_wide and bays_wide > 0:
        # Use custom bay widths if provided, otherwise divide equally
        if custom_bay_widths and len(custom_bay_widths) == bays_wide:
            # Convert custom bay widths from inches to feet
            bay_widths_ft = [w / 12.0 for w in custom_bay_widths]
        else:
            # Equal division
            bay_width_ft = (opening_width / bays_wide) / 12
            bay_widths_ft = [bay_width_ft] * bays_wide

        if total_count > 1:
            return bay_widths_ft * total_count
        return bay_widths_ft

    single_instance_qty = opening_width / 12
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty


def calculate_og_head_h(
    opening_width: float,
    total_count: int,
    bays_wide: int = None,
    custom_bay_widths: list = None,
) -> Union[float, list[float]]:
    """
    Calculates outside glazing head horizontal feet. Returns a list of bay widths if bays_wide is provided,
    otherwise returns the whole width. Uses custom_bay_widths if provided, otherwise divides equally.
    Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2514 (uses bay widths)
    """
    if bays_wide and bays_wide > 0:
        # Use custom bay widths if provided, otherwise divide equally
        if custom_bay_widths and len(custom_bay_widths) == bays_wide:
            # Convert custom bay widths from inches to feet
            bay_widths_ft = [w / 12.0 for w in custom_bay_widths]
        else:
            # Equal division
            bay_width_ft = (opening_width / bays_wide) / 12
            bay_widths_ft = [bay_width_ft] * bays_wide

        if total_count > 1:
            return bay_widths_ft * total_count
        return bay_widths_ft

    single_instance_qty = opening_width / 12
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty


def calculate_sill_flashing_h(
    opening_width: float, total_count: int
) -> Union[float, list[float]]:
    """
    Calculates sill flashing horizontal feet. Returns a list if total_count > 1, else a float.
    Associated with profile: BE9-2578
    """
    single_instance_qty = opening_width / 12
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty


def calculate_fabrication_joints(
    bays_wide: int, bays_tall: int, total_count: int
) -> int:
    """Calculate number of fabrication joints."""
    return ((4 * bays_wide) + (bays_wide * (2 * (bays_tall - 1)))) * total_count


def calculate_glass_stop(
    opening_width: float,
    bays_tall: int,
    total_count: int,
    bays_wide: int = None,
    custom_bay_widths: list = None,
) -> Union[float, list[float]]:
    """
    Calculate glass stop length. Returns a list of bay widths repeated for each bay (total bays = bays_wide * bays_tall).
    Uses custom_bay_widths if provided, otherwise divides equally.
    Returns a list if total_count > 1, else a float.
    Associated with profile: E9-2519 (uses bay widths, one per bay)
    """
    if bays_wide and bays_wide > 0:
        # Use custom bay widths if provided, otherwise divide equally
        if custom_bay_widths and len(custom_bay_widths) == bays_wide:
            # Convert custom bay widths from inches to feet
            bay_widths_ft = [w / 12.0 for w in custom_bay_widths]
        else:
            # Equal division
            bay_width_ft = (opening_width / bays_wide) / 12
            bay_widths_ft = [bay_width_ft] * bays_wide

        # Total number of bays = bays_wide * bays_tall
        # Each bay gets one glass stop piece, so repeat bay_widths_ft for each row (bays_tall)
        total_bays = bays_wide * bays_tall
        bay_widths_repeated = bay_widths_ft * bays_tall

        if total_count > 1:
            return bay_widths_repeated * total_count
        return bay_widths_repeated

    single_instance_qty = (opening_width / 12) * bays_tall
    if total_count > 1:
        return [single_instance_qty] * total_count
    return single_instance_qty


# ================================================================
# D.L.O. (Daylight Opening) & Glass Make Size
# YES 45 TU framing deductions:
#   Edge bay  (jamb/head/sill to intermediate) : -3"
#   Interior  (intermediate to intermediate)    : -2"
#   Bottom sill additional                      : -7/16"
#   Glass make size = D.L.O. + 3/4"
# ================================================================

DLO_EDGE_DEDUCTION = 3  # inches — jamb/head/sill to intermediate
DLO_INTERIOR_DEDUCTION = 2  # inches — intermediate to intermediate
DLO_SILL_DEDUCTION = 7 / 16  # inches — additional at bottom of elevation
GLASS_MAKE_ADDITION = 0.75  # inches — glass make = DLO + 3/4"


def _dlo_width(cl_width: float, col_index: int, total_cols: int) -> float:
    is_edge = total_cols <= 1 or col_index == 0 or col_index == total_cols - 1
    return cl_width - (DLO_EDGE_DEDUCTION if is_edge else DLO_INTERIOR_DEDUCTION)


def _dlo_height(cl_height: float, row_index: int, total_rows: int) -> float:
    is_edge = total_rows <= 1 or row_index == 0 or row_index == total_rows - 1
    is_bottom = row_index == 0
    dlo = cl_height - (DLO_EDGE_DEDUCTION if is_edge else DLO_INTERIOR_DEDUCTION)
    if is_bottom:
        dlo -= DLO_SILL_DEDUCTION
    return dlo


def calculate_total_glass(
    opening_width: float,
    opening_height: float,
    total_count: int,
    bays_wide: int,
    bays_tall: int,
    custom_bay_widths=None,
    custom_bay_heights=None,
) -> float:
    """Calculate total glass area (sqft) using per-lite glass make sizes (DLO + 3/4\")."""
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
        dlo_w = _dlo_width(bay_widths[col], col, bays_wide)
        glass_w = dlo_w + GLASS_MAKE_ADDITION
        for row in range(bays_tall):
            dlo_h = _dlo_height(bay_heights[row], row, bays_tall)
            glass_h = dlo_h + GLASS_MAKE_ADDITION
            total_sqft += (glass_w * glass_h) / 144

    return total_sqft * total_count


def calculate_door_size(door_size_str: str) -> float:
    """
    Calculates the area of a single door from a size string like "3' X 7'".
    Returns the area in square feet.
    """
    try:
        # Normalize the string and split on 'X'
        parts = door_size_str.upper().replace(" ", "").split("X")
        if len(parts) != 2:
            raise ValueError(f"Invalid door size format: {door_size_str}")

        # Remove the apostrophe and convert to float
        width_ft = float(parts[0].replace("'", ""))
        height_ft = float(parts[1].replace("'", ""))

        area = width_ft * height_ft
        return area

    except Exception as e:
        print(f"Error calculating door area for '{door_size_str}': {e}")
        return 0.0


def calculate_door_price(
    size_str: str, width_type: str, hardware_dict: dict, finish: str
) -> float:
    """
    Calculates total price of a door given size (e.g. "3' X 8'"), width_type ("Narrow", "Medium", "Wide"),
    finish ("Clear", "Black", "Paint"), and selected hardware.
    Hardware prices also vary by finish.
    """

    # Full door price matrix
    DOOR_PRICES = {
        "3x7": {
            "Narrow": {"Clear": 880.00, "Black": 1035.00, "Paint": 1269.00},
            "Medium": {"Clear": 1180.00, "Black": 1245.00, "Paint": 1653.00},
            "Wide": {"Clear": 1304.25, "Black": 1413.75, "Paint": 1744.50},
        },
        "3x8": {
            "Narrow": {"Clear": 921.75, "Black": 1083.00, "Paint": 1328.25},
            "Medium": {"Clear": 1235.25, "Black": 1304.25, "Paint": 1727.25},
            "Wide": {"Clear": 1365.00, "Black": 1479.00, "Paint": 1825.50},
        },
        "3x9": {
            "Narrow": {"Clear": 986.25, "Black": 1159.50, "Paint": 1422.75},
            "Medium": {"Clear": 1321.50, "Black": 1395.75, "Paint": 1849.50},
            "Wide": {"Clear": 1461.75, "Black": 1584.00, "Paint": 1953.00},
        },
        "6x7": {
            "Narrow": {"Clear": 1715.25, "Black": 1863.75, "Paint": 2657.25},
            "Medium": {"Clear": 2310.75, "Black": 2445.75, "Paint": 3156.75},
            "Wide": {"Clear": 2559.00, "Black": 2781.00, "Paint": 3435.75},
        },
        "6x8": {
            "Narrow": {"Clear": 1812.00, "Black": 1970.25, "Paint": 2799.75},
            "Medium": {"Clear": 2442.00, "Black": 2589.75, "Paint": 3338.25},
            "Wide": {"Clear": 2700.00, "Black": 2932.50, "Paint": 3630.00},
        },
        "6x9": {
            "Narrow": {"Clear": 1943.25, "Black": 2111.25, "Paint": 2988.00},
            "Medium": {"Clear": 2624.25, "Black": 2782.50, "Paint": 3564.00},
            "Wide": {"Clear": 2901.00, "Black": 3150.00, "Paint": 3861.00},
        },
    }

    # Hardware base prices by finish - Updated keys to match UI options
    HARDWARE_PRICES = {
        "Concealed Closer": {"Clear": 473.00, "Black": 473.00, "Paint": 473.00},
        "Exit Devices": {
            "Clear": 475.00,
            "Black": 475.00,
            "Paint": 475.00,
        },  # "Exit Devices" in UI
        "Exit Device": {
            "Clear": 475.00,
            "Black": 475.00,
            "Paint": 475.00,
        },  # Legacy key
        "Continuous Hinges": {"Clear": 285.00, "Black": 375.00, "Paint": 375.00},
        "Latch Lock w/ Lever Handle": {
            "Clear": 334.00,
            "Black": 334.00,
            "Paint": 334.00,
        },  # "Latch Lock w/ Lever Handle" in UI
        "Lever Handle": {
            "Clear": 167.00,
            "Black": 167.00,
            "Paint": 167.00,
        },  # Estimation: Half of Latch+Lever set
        "Electric Strike": {"Clear": 355.00, "Black": 355.00, "Paint": 355.00},
        "Extended Ladder Pull (B2B)": {
            "Clear": 350.00,
            "Black": 400.00,
            "Paint": 400.00,
        },  # Placeholder price
        "Extended Ladder Pull (Single)": {
            "Clear": 175.00,
            "Black": 200.00,
            "Paint": 200.00,
        },  # Placeholder price
    }

    try:
        # Normalize size key (e.g., "3' X 8'" -> "3x8")
        parts = size_str.upper().replace("'", "").replace(" ", "").split("X")
        if len(parts) != 2:
            raise ValueError(f"Invalid door size format: {size_str}")
        door_size_key = f"{parts[0].lower()}x{parts[1].lower()}"

        # Normalize finish to Title Case (e.g. "clear" -> "Clear") to match keys
        finish_key = finish.title()
        if finish_key not in ["Clear", "Black", "Paint"]:
            finish_key = "Clear"  # Default to Clear if unknown

        # Get base price
        base_price = DOOR_PRICES[door_size_key][width_type][finish_key]

        # Calculate hardware cost based on finish
        hardware_total = 0
        for hw, selected in hardware_dict.items():
            if selected and hw in HARDWARE_PRICES:
                price = HARDWARE_PRICES[hw][finish_key]

                # Special rule: Exit Device is $550 for all doors except 3x7 and 6x7
                if (
                    hw == "Exit Device" or hw == "Exit Devices"
                ) and door_size_key not in ["3x7", "6x7"]:
                    price = 550.00

                # Double price for double doors (all hardware)
                if door_size_key.startswith("6x"):
                    price *= 2

                hardware_total += price

        return base_price + hardware_total

    except KeyError as e:
        print(f"Invalid key in pricing lookup: {e}")
        return 0.0
    except Exception as e:
        print(f"Error calculating price: {e}")
        return 0.0


def calculate_door_info(doors: list, finish="Clear", total_count: int = 1) -> list:
    """
    Takes a list of door inputs and returns a list of dictionaries with door information,
    including calculated price and other details.

    Note: door count is per elevation, so it's multiplied by total_count to get total quantity.

    Args:
        doors (list): A list of dictionaries, where each dictionary represents a door
                      and contains keys like 'size', 'count' (per elevation), 'stile', and 'hardware'.
        finish (str): Finish type for doors
        total_count (int): Total number of elevations. Door count is per elevation, so multiply by this.

    Returns:
        list: A list of dictionaries, where each dictionary represents a door with
              its calculated details and is formatted for the final output.
    """
    door_items = []
    if doors:
        for door_info in doors:
            door_size_str = door_info.get("size")
            door_count_per_elev = door_info.get("count", 0)  # Count is per elevation
            door_stile = door_info.get("stile")
            door_hardware = door_info.get("hardware", {})

            if door_size_str and door_count_per_elev > 0:
                # Calculate total door count (per elevation * total_count)
                total_door_count = door_count_per_elev * total_count
                door_price = calculate_door_price(
                    door_size_str, door_stile, door_hardware, finish
                )

                door_items.append(
                    {
                        "description": f"Door ({door_size_str})",
                        "Style": door_stile,
                        "quantity": total_door_count,  # Total quantity across all elevations
                        "part_number": "N/A",
                        "type": "Doors",
                        "price": door_price,
                        "hardware": door_hardware,
                        "manual": True,
                    }
                )
    return door_items


def calculate_total_door_area(doors: list) -> float:
    """
    Calculates the total area (in sqft) of all doors in the list.
    Args:
        doors (list): List of door dicts, each with 'size' and 'count'.
    Returns:
        float: Total area in sqft.
    """
    total_area = 0.0
    for door in doors:
        size_str = door.get("size")
        count = door.get("count", 0)
        if size_str and count:
            area = calculate_door_size(size_str)
            total_area += area * count
    return total_area


def validate_door_fits(door, total_glass_area: float) -> dict:
    """
    Validate if a door opening would make glass area go negative.

    Args:
        door (dict): Door with 'size', 'count', 'stile' keys.
        total_glass_area (float): Total glass area in sqft (from calculate_total_glass).

    Returns:
        dict: {'valid': bool, 'message': str, 'glass_after': float}
    """
    deductions = {
        "Narrow": {"height": 13.5625, "width": 4.875},
        "Medium": {"height": 15.1875, "width": 7.625},
        "Wide": {"height": 16.25, "width": 10.625},
    }

    stile = door.get("stile", "").title()
    if stile not in deductions:
        return {"valid": False, "message": "Invalid stile type", "glass_after": total_glass_area}

    size_str = door.get("size", "")
    if not size_str:
        return {"valid": False, "message": "No door size specified", "glass_after": total_glass_area}

    # Parse door size
    size_match = re.match(r"\s*(\d+)' *[xX] *(\d+)'", size_str)
    if not size_match:
        return {"valid": False, "message": f"Invalid door size format: {size_str}", "glass_after": total_glass_area}

    opening_width_ft = int(size_match.group(1))
    opening_height_ft = int(size_match.group(2))
    opening_width_in = opening_width_ft * 12
    opening_height_in = opening_height_ft * 12

    # Calculate door glass area
    if opening_width_ft == 6:
        glass_width = (opening_width_in / 2) - deductions[stile]["width"]
    else:
        glass_width = opening_width_in - deductions[stile]["width"]

    glass_height = opening_height_in - deductions[stile]["height"]

    if glass_width <= 0 or glass_height <= 0:
        return {
            "valid": False,
            "message": f"Door stile '{stile}' too large for {size_str} opening",
            "glass_after": total_glass_area
        }

    door_count = door.get("count", 1)
    door_glass_area = (glass_width * glass_height) / 144 * door_count
    glass_after = round(total_glass_area - door_glass_area, 2)

    if glass_after < 0:
        return {
            "valid": False,
            "message": f"⚠️ CAUTION: Door opening ({door_glass_area:.2f} sqft) exceeds available glass ({total_glass_area:.2f} sqft). Result: {glass_after:.2f} sqft (NEGATIVE)",
            "glass_after": glass_after,
            "door_area": door_glass_area
        }
    elif glass_after < 0.5:
        return {
            "valid": True,
            "message": f"⚠️ WARNING: Very little glass remaining ({glass_after:.2f} sqft). Door opening: {door_glass_area:.2f} sqft",
            "glass_after": glass_after,
            "door_area": door_glass_area,
            "is_warning": True
        }

    return {
        "valid": True,
        "message": f"Door fits. Glass remaining: {glass_after:.2f} sqft",
        "glass_after": glass_after,
        "door_area": door_glass_area
    }


def calculate_glass_to_add_back(doors):
    """
    Calculate total glass back area in sqft based on door sizes.

    Args:
        doors (list of dict): List of doors with 'size', 'count', 'stile' keys.

    Returns:
        float: Total glass back area in sqft.
    """

    deductions = {
        "Narrow": {"height": 13.5625, "width": 4.875},
        "Medium": {"height": 15.1875, "width": 7.625},
        "Wide": {"height": 16.25, "width": 10.625},
    }

    if not doors or not isinstance(doors, list):
        return 0

    total_area = 0.0

    for door in doors:
        stile = door.get("stile", "").title()
        if stile not in deductions:
            continue

        count = door.get("count", 1)
        size_str = door.get("size", "")
        if not size_str:
            continue

        # Parse door opening width and height in feet from 'size' string
        size_match = re.match(r"\s*(\d+)' *[xX] *(\d+)'", size_str)
        if not size_match:
            continue

        opening_width_ft = int(size_match.group(1))
        opening_height_ft = int(size_match.group(2))

        # Convert to inches
        opening_width_in = opening_width_ft * 12
        opening_height_in = opening_height_ft * 12

        # For 6' width doors, width deduction applies after dividing width by 2 (paired door)
        if opening_width_ft == 6:
            glass_width = (opening_width_in / 2) - deductions[stile]["width"]
        else:
            glass_width = opening_width_in - deductions[stile]["width"]

        glass_height = opening_height_in - deductions[stile]["height"]

        if glass_width <= 0 or glass_height <= 0:
            continue

        area_sqft = (glass_width * glass_height) / 144
        total_area += area_sqft * count

    return round(total_area, 2)
