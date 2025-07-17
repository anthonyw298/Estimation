import pandas as pd
import re

FILE_PATH = r'C:\Users\tonyw\OneDrive\Desktop\Estimation\data\mydata.xlsx'
OUTPUT_CSV_PATH = r'C:\Users\tonyw\OneDrive\Desktop\Estimation\data\cleaned_data.csv'  # changed output path

def clean_value(val):
    if isinstance(val, str):
        return val.strip()
    return val

def extract_fields(row):
    part_number = None
    finish = None
    length = None
    units = None
    list_price = None

    found_idx = -1
    for idx in range(len(row)):
        val = row.iloc[idx]
        if pd.notna(val):
            s = str(val).strip()
            if re.match(r'^[A-Za-z0-9\s]*-\d+', s):
                part_number = s
                found_idx = idx
                break

    if found_idx == -1:
        return pd.Series([None, None, None, None, None, None])

    for idx in range(found_idx + 1, len(row)):
        val = row.iloc[idx]
        if pd.isna(val):
            continue
        s = str(val).strip()
        if s == '':
            continue

        if finish is None and re.match(r'^[A-Za-z0-9\s\(\)]+$', s):
            finish = s
            continue

        if length is None and re.search(r"[\"\'\u2018\u2019\u201C\u201D]", s):
            length = s
            continue

        if units is None:
            s_lower = s.lower()
            if 'pc' in s_lower or 'pcs' in s_lower:
                units = s
                continue

        if list_price is None:
            try:
                num = float(s.replace('$', '').replace(',', '').strip())
                list_price = num
                continue
            except ValueError:
                pass

    page_number = row.iloc[-1] if len(row) > 0 else None

    return pd.Series([part_number, finish, length, units, list_price, page_number])

def main():
    USE_COLS = list(range(16))

    df = pd.read_excel(FILE_PATH, header=None, usecols=USE_COLS, skiprows=2)
    df = df.apply(lambda col: col.map(clean_value))
    extracted = df.apply(extract_fields, axis=1)
    extracted.columns = ['Part Number', 'Finish', 'Length', 'Units', 'List Price', 'Page Number(s)']

    extracted = extracted[extracted['Part Number'].notna()]
    extracted['List Price'] = pd.to_numeric(extracted['List Price'], errors='coerce')
    extracted.index = extracted.index + 3

    # Save to CSV file instead of .py
    extracted.to_csv(OUTPUT_CSV_PATH, index=False)

    print(f"Saved cleaned data to CSV file at {OUTPUT_CSV_PATH}")

if __name__ == "__main__":
    main()
