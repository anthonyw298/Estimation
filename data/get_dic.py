import csv

import csv
import json

def build_parts_dictionary(csv_filename):
    parts_dict = {}

    with open(csv_filename, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)

        for row in reader:
            part_number = row["Part Number"].strip()
            finish = row["Finish"].strip()
            length = row["Length"].strip()
            units = row["Units"].strip()

            list_price_str = row["List Price"].strip()
            if list_price_str:
                list_price = float(list_price_str)
            else:
                list_price = 0.0  # fallback if missing

            page_numbers = [p.strip() for p in row["Page Number(s)"].split(",")]

            if part_number not in parts_dict:
                parts_dict[part_number] = {
                    "Finish": [finish],
                    "Length": length,
                    "Units": units,
                    "List Price": list_price,
                    "Page Numbers": page_numbers
                }
            else:
                if finish not in parts_dict[part_number]["Finish"]:
                    parts_dict[part_number]["Finish"].append(finish)
                if list_price < parts_dict[part_number]["List Price"]:
                    parts_dict[part_number]["List Price"] = list_price
                for pn in page_numbers:
                    if pn not in parts_dict[part_number]["Page Numbers"]:
                        parts_dict[part_number]["Page Numbers"].append(pn)

    return parts_dict

# ✅ Build the dictionary
parts_data = build_parts_dictionary(r"C:\Users\tonyw\OneDrive\Desktop\Estimation\data\cleaned_data.csv")

# ✅ Write it as a new .py file
with open(r"C:\Users\tonyw\OneDrive\Desktop\Estimation\data\parts_data.py", "w", encoding="utf-8") as pyfile:
    pyfile.write("parts_data = ")
    pyfile.write(json.dumps(parts_data, indent=4))
