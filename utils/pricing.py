import pandas as pd

def get_price_by_part(part_number, csv_file='cleaned_data.csv'):
    df = pd.read_csv(csv_file)
    #match = df[df['Part Number'] == part_number]
    match = df[df['Part Number'] == part_number]
    if match.empty:
        return None
    print(match['List Price'].values[0])
    return match['List Price'].values[0]


##DEFAULTS THE FINISH