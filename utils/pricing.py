import os
import pandas as pd

def get_price_by_part(part_number):
    # Always find the data file relative to this script's folder
    base_dir = os.path.dirname(__file__)  # This is utils/
    project_root = os.path.dirname(base_dir)  # This goes up one level to Estimation/
    csv_path = os.path.join(project_root, 'data', 'cleaned_data.csv')
    df = pd.read_csv(csv_path)

    match = df[df['Part Number'] == part_number]
    if match.empty:
        return None
    return match['List Price'].values[0]
