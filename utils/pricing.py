import re
import json
import os
from data.parts_data import parts_data
from data.part_number import PART_NUMBER_MAP

EXTRA_MATERIALS_FILE = "extra_materials.json"

def parse_length_to_feet(length_str):
    """
    Converts various length formats to total feet.
    Examples: 8', 96", 8 ft, 8ft 6in.
    Returns 0.0 if input is invalid or empty.
    """
    if not isinstance(length_str, str) or not length_str.strip():
        return 0.0

    length_str = length_str.replace('’', "'").replace('”', '"').replace('“', '"')

    feet = 0.0
    inches = 0.0

    feet_match = re.search(r"(\d+\.?\d*)\s*(ft|')", length_str, re.IGNORECASE)
    if feet_match:
        feet = float(feet_match.group(1))

    inches_match = re.search(r"(\d+\.?\d*)\s*(in|\")", length_str, re.IGNORECASE)
    if inches_match:
        inches = float(inches_match.group(1))

    if feet or inches:
        return feet + (inches / 12)

    num_match = re.search(r"(\d+\.?\d*)", length_str)
    if num_match:
        return float(num_match.group(1))

    return 0.0

def load_extra_materials():
    """Load extra materials leftovers from JSON file."""
    if os.path.exists(EXTRA_MATERIALS_FILE):
        try:
            with open(EXTRA_MATERIALS_FILE, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            print(f"Warning: Could not decode {EXTRA_MATERIALS_FILE}: {e}. Starting with empty extra materials.")
            return {}
    return {}

def save_extra_materials(materials):
    """Save extra materials leftovers to JSON file."""
    try:
        with open(EXTRA_MATERIALS_FILE, 'w') as f:
            json.dump(materials, f, indent=4)
    except IOError as e:
        print(f"Error: Could not save {EXTRA_MATERIALS_FILE}: {e}")

def get_price_by_part(part_number, requested_qty):
    """
    Calculate price using:
      - profiles_group: length-based
      - accessories_group: piece-based
    Uses leftover materials if possible and saves updated leftovers.
    """
    match = parts_data.get(part_number)
    if not match:
        return None, None, 0, 0.0

    list_price = float(match.get('List Price', 0.0))
    units_str = match.get('Units', "1 pcs.")
    length_str = match.get('Length', "")

    extra_materials = load_extra_materials()
    part_extra = extra_materials.get(part_number, {'quantity': 0, 'length_pieces': []})

    total_price = 0.0
    actual_purchased_qty = 0
    actual_purchased_length = 0.0

    if part_number in PART_NUMBER_MAP['profiles']:
        # Length-based
        unit_type = "ft"
        min_purchase_length = parse_length_to_feet(length_str)
        if min_purchase_length <= 0:
            min_purchase_length = 1.0

        leftover_pieces = part_extra.get('length_pieces', [])

        # Sort leftover pieces smallest to largest to always use smallest fitting piece
        leftover_pieces.sort()

        suitable_index = None
        for i, piece_len in enumerate(leftover_pieces):
            if piece_len >= requested_qty:
                suitable_index = i
                break

        if suitable_index is not None:
            used_piece = leftover_pieces.pop(suitable_index)
            leftover_after_use = used_piece - requested_qty
            if leftover_after_use > 0:
                leftover_pieces.append(leftover_after_use)

            total_price = 0.0  # Used leftover only
            actual_purchased_length = 0.0
        else:
            num_bundles_needed = int(-(-requested_qty // min_purchase_length))  # ceil div
            actual_purchased_length = num_bundles_needed * min_purchase_length
            total_price = list_price * num_bundles_needed

            leftover_piece = actual_purchased_length - requested_qty
            if leftover_piece > 0:
                leftover_pieces.append(leftover_piece)

        part_extra['length_pieces'] = leftover_pieces
        part_extra['quantity'] = 0

    elif part_number in PART_NUMBER_MAP['accessories']:
        # Piece-based
        unit_type = "pcs"
        unit_count_per_bundle = 1
        if 'pc' in units_str.lower():
            try:
                pcs_part = units_str.lower().split('pc')[0].strip()
                unit_count_per_bundle = int(pcs_part) if pcs_part else 1
            except ValueError:
                pass

        leftover_qty = part_extra.get('quantity', 0)
        remaining_needed_qty = requested_qty - leftover_qty

        if remaining_needed_qty <= 0:
            actual_purchased_qty = 0
            total_price = 0.0
            excess_qty_after_use = leftover_qty - requested_qty
        else:
            num_bundles_needed = (remaining_needed_qty + unit_count_per_bundle - 1) // unit_count_per_bundle
            actual_purchased_qty = num_bundles_needed * unit_count_per_bundle
            total_price = list_price * num_bundles_needed
            excess_qty_after_use = leftover_qty + actual_purchased_qty - requested_qty

        part_extra['quantity'] = excess_qty_after_use
        part_extra['length_pieces'] = []

    else:
        # Not found in either group
        return None, None, 0, 0.0

    extra_materials[part_number] = part_extra
    save_extra_materials(extra_materials)

    return total_price, unit_type
