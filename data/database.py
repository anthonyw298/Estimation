import pandas as pd
import re
import json

FILE_PATH = r'C:\Users\tonyw\OneDrive\Desktop\Estimation\data\mydata.xlsx'
OUTPUT_DICT_PATH = r'C:\Users\tonyw\OneDrive\Desktop\Estimation\data\cleaned_data.py'  # saving as a .py file

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

    # Build dictionary output
    output_dict = {}
    for _, row in extracted.iterrows():
        pn = row['Part Number']
        if pn is None:
            continue

        # Finish as list
        finish_list = []
        if isinstance(row['Finish'], str):
            finish_list = [f.strip() for f in row['Finish'].split(',') if f.strip()]
        # Page Numbers as list, split by comma or keep as list if already
        pages = []
        if pd.notna(row['Page Number(s)']):
            if isinstance(row['Page Number(s)'], str):
                pages = [p.strip() for p in row['Page Number(s)'].split(',') if p.strip()]
            elif isinstance(row['Page Number(s)'], list):
                pages = row['Page Number(s)']
            else:
                pages = [str(row['Page Number(s)'])]

        output_dict[pn] = {
            "Finish": finish_list if finish_list else [],
            "Length": row['Length'] if pd.notna(row['Length']) else "",
            "Units": row['Units'] if pd.notna(row['Units']) else "",
            "List Price": float(row['List Price']) if pd.notna(row['List Price']) else None,
            "Page Numbers": pages
        }

    # Save dictionary as a Python file with valid syntax
    with open(OUTPUT_DICT_PATH, 'w', encoding='utf-8') as f:
        f.write("data = ")
        f.write(json.dumps(output_dict, indent=4, ensure_ascii=False))

    print(f"Saved cleaned data dictionary to {OUTPUT_DICT_PATH}")

if __name__ == "__main__":
    main()
