import os
import pandas as pd
import re

def parse_length_to_feet(length_str):
    """
    Converts various length formats to total feet.
    Handles:
    - 10'
    - 10' - 6"
    - 10'-6"
    - 10 FT 6 IN
    - 120 IN
    - 120"
    - 10’ 6”
    - 5.5'
    - 66"
    """
    if not isinstance(length_str, str):
        return 1

    # Normalize curly quotes
    length_str = length_str.replace('’', "'").replace('”', '"').replace('“', '"')

    feet = 0.0
    inches = 0.0

    # Match feet (e.g., 10', 10 ft, 5.5')
    feet_match = re.search(r"(\d+\.?\d*)\s*(ft|')", length_str, re.IGNORECASE)
    if feet_match:
        feet = float(feet_match.group(1))

    # Match inches (e.g., 6", 6 in)
    inches_match = re.search(r"(\d+\.?\d*)\s*(in|\")", length_str, re.IGNORECASE)
    if inches_match:
        inches = float(inches_match.group(1))

    # Only inches
    if not feet and inches:
        return inches / 12

    # Only feet
    if feet and not inches:
        return feet

    # Both
    if feet or inches:
        return feet + (inches / 12)

    return 1  # fallback

def get_price_by_part(part_number):
    # Always find the data file relative to this script's folder
    base_dir = os.path.dirname(__file__)
    project_root = os.path.dirname(base_dir)
    csv_path = os.path.join(project_root, 'data', 'cleaned_data.csv')
    df = pd.read_csv(csv_path)

    match = df[df['Part Number'] == part_number]
    if match.empty:
        return None

    # Get List Price
    list_price = match['List Price'].values[0]

    # --------------------
    # Adjust for Units (pcs., pc., Roll)
    # --------------------
    units_str = match['Units'].values[0] if 'Units' in match.columns else None
    unit_count = 1  # default

    if isinstance(units_str, str):
        units_lower = units_str.lower().strip()
        if 'pcs' in units_lower or 'pc' in units_lower:
            try:
                unit_count = int(units_lower.split('pc')[0].strip())
            except Exception:
                unit_count = 1
    if unit_count > 1:
        list_price /= unit_count

    # --------------------
    # Adjust for Length
    # --------------------
    length_str = match['Length'].values[0] if 'Length' in match.columns else None
    length_ft = parse_length_to_feet(length_str)

    if length_ft > 1:
        list_price /= length_ft

    return list_price
