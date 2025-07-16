import pandas as pd
import re

FILE_PATH = r'C:\Users\tonyw\OneDrive\Desktop\Estimation\data\mydata.xlsx'

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
    for idx in range(16):
        val = row.iloc[idx]
        if pd.notna(val):
            s = str(val).strip()
            if re.match(r'^[A-Za-z0-9\s]*-\d+', s):
                part_number = s
                found_idx = idx
                break

    if found_idx == -1:
        return pd.Series([None, None, None, None, None, None])

    for idx in range(found_idx + 1, 16):
        val = row.iloc[idx]
        if pd.isna(val) or val == '':
            continue
        s = str(val).strip()

        # ✅ NEW: allow numbers and parentheses
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
            if '$' in s or re.match(r'^\d+(\.\d+)?$', s):
                list_price = s
                continue

    page_number = row.iloc[15] if len(row) > 15 else None

    return pd.Series([part_number, finish, length, units, list_price, page_number])


def main():
    filter_prefix = 'AS-'               #CHANGE TO WHAT UR SEARCHING FOR
    USE_COLS = list(range(16))

    df = pd.read_excel(FILE_PATH, header=None, usecols=USE_COLS, skiprows=2)

    # ✅ Use .apply instead of deprecated .applymap
    df = df.apply(lambda col: col.map(clean_value))

    extracted = df.apply(extract_fields, axis=1)
    extracted.columns = ['Part Number', 'Finish', 'Length', 'Units', 'List Price', 'Page Number(s)']

    extracted = extracted[
        extracted['Part Number'].notna() & extracted['Part Number'].str.startswith(filter_prefix)
    ]

    # ✅ Use raw string for regex to avoid SyntaxWarning
    extracted['List Price'] = extracted['List Price'].replace(r'[\$,]', '', regex=True)
    extracted['List Price'] = pd.to_numeric(extracted['List Price'], errors='coerce')

    # Adjust index to match skipped rows
    extracted.index = extracted.index + 3


if __name__ == "__main__":
    main()
