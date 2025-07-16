import pandas as pd

def get_price_by_part(part_number, csv_file='cleaned_data.csv'):
    # Load the cleaned data CSV
    df = pd.read_csv(csv_file)
    
    # Filter dataframe for matching part number
    match = df[df['Part Number'] == part_number]
    
    if match.empty:
        return None  # No match found
    
    # Return the first matching List Price
    return match['List Price'].values[0]

if __name__ == "__main__":
    # Change this part number to whatever you want to look up
    part_to_lookup = 'E2-0052'  
    
    price = get_price_by_part(part_to_lookup)
    
    if price is not None:
        print(f"Price for {part_to_lookup} is {price}")
    else:
        print(f"Part number {part_to_lookup} not found.")
