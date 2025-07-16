#df = pd.read_excel(r'C:\Users\tonyw\OneDrive\Desktop\Estimation\data\mydata.xlsx')
import pandas as pd

# Columns to read from Excel: A, C, G, M, N, O (zero-based indices)
USE_COLS = [0, 2, 6, 12, 13, 15]

# Column names to assign after reading
COLUMNS = ['Part Number', 'Finish', 'Length', 'Units', 'List Price', 'Page Number(s)']

# Path to your Excel file
FILE_PATH = r'C:\Users\tonyw\OneDrive\Desktop\Estimation\data\mydata.xlsx'

def clean_value(val):
    # Strip whitespace if string
    return val.strip() if isinstance(val, str) else val

def main():
    # Read specified columns, skip the first row (header row)
    df = pd.read_excel(FILE_PATH, header=None, usecols=USE_COLS, skiprows=1)
    
    # Clean whitespace in all cells
    df_cleaned = df.applymap(clean_value)
    
    # Assign column names
    df_cleaned.columns = COLUMNS

    # Drop rows with empty or NaN part numbers
    df_cleaned = df_cleaned[df_cleaned['Part Number'].notna() & (df_cleaned['Part Number'] != '')]
    
    # Keep only rows where Part Number starts with 'AS-'
    df_cleaned = df_cleaned[df_cleaned['Part Number'].str.startswith('AS-')]

    # Save cleaned data to CSV for fetch.py to use
    df_cleaned.to_csv('cleaned_data.csv', index=False)

if __name__ == "__main__":
    main()
