import re

def parse_length_to_feet(length_str):
    """
    Converts various length formats to total feet.
    Accepts formats like: 8', 96", 8 ft, 8ft 6in, etc.
    """
    if not isinstance(length_str, str) or not length_str.strip():
        return 1

    # Normalize smart quotes to plain
    length_str = length_str.replace('’', "'").replace('”', '"').replace('“', '"')

    feet = 0.0
    inches = 0.0

    # Match feet
    feet_match = re.search(r"(\d+\.?\d*)\s*(ft|')", length_str, re.IGNORECASE)
    if feet_match:
        feet = float(feet_match.group(1))

    # Match inches
    inches_match = re.search(r"(\d+\.?\d*)\s*(in|\")", length_str, re.IGNORECASE)
    if inches_match:
        inches = float(inches_match.group(1))

    if feet or inches:
        return feet + (inches / 12)

    # Fallback if nothing matches
    return 1


# ✅ Now for dict-based parts_data:
from data.parts_data import parts_data  # assuming you saved parts_data.py

def get_price_by_part(part_number):
    """
    Gets adjusted price per piece per foot for given part number.
    Works with dict-based parts_data.
    """
    match = parts_data.get(part_number)
    if not match:
        return None

    list_price = match.get('List Price', 0)
    units_str = match.get('Units', None)
    unit_count = 1

    if isinstance(units_str, str):
        units_lower = units_str.lower().strip()
        if 'pcs' in units_lower or 'pc' in units_lower:
            try:
                unit_count = int(units_lower.split('pc')[0].strip())
            except Exception:
                unit_count = 1

    if unit_count > 1:
        list_price /= unit_count

    length_str = match.get('Length', None)
    length_ft = parse_length_to_feet(length_str)

    if length_ft > 1:
        list_price /= length_ft

    return list_price
