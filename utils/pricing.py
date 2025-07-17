import os
import pandas as pd

def get_price_by_part(part_number):
    # Always find the data file relative to this script's folder
    base_dir = os.path.dirname(__file__)  # This is utils/
    project_root = os.path.dirname(base_dir)  # Go up one level to Estimation/
    csv_path = os.path.join(project_root, 'data', 'cleaned_data.csv')
    df = pd.read_csv(csv_path)

    match = df[df['Part Number'] == part_number]
    if match.empty:
        return None

    list_price = match['List Price'].values[0]
    units_val = match['Units'].values[0]

    # Try to get number of units from Units field
    try:
        # If units_val is already numeric (int/float)
        units_num = int(units_val)
    except (ValueError, TypeError):
        # Otherwise parse string like '100 pcs.'
        try:
            units_str = str(units_val)
            units_num = int(units_str.split()[0])
        except (IndexError, ValueError):
            units_num = 1

    # Calculate unit price if bulk quantity > 1
    if units_num > 1:
        unit_price = float(list_price) / units_num
    else:
        unit_price = float(list_price)

    return unit_price

