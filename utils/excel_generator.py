import openpyxl
from openpyxl.utils import get_column_letter
from data.part_number import PART_NUMBER_MAP
from utils.pricing import get_price_by_part

def generate_excel_report(
    system_input: str,
    finish_input: str,
    elevation_type: str,
    total_count: int,
    bays_wide: int,
    bays_tall: int,
    opening_width: float,
    opening_height: float,
    sqft_per_type: float,
    total_sqft: float,
    perimeter_ft: float,
    total_perimeter_ft: float,
    calculated_outputs: list,
    completion_callback=None
):
    if system_input != "YES 45TU FRONT SET(OG)":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = f"System '{system_input}' not matched. Empty file created."
        wb.save("output.xlsx")
        if completion_callback:
            completion_callback("System not matched. Empty 'output.xlsx' created.", "orange")
        return

    finish_multiplier_map = {
        "clear": 1.0,
        "black": 1.1,
        "paint": 1.2
    }
    multiplier = finish_multiplier_map.get(finish_input.lower(), 1.0)

    wb = openpyxl.Workbook()
    ws = wb.active

    inputs_headers = [
        "System Input", "Elevation Type", "Total Count",
        "# Bays Wide", "# Bays Tall",
        "Opening Width", "Opening Height",
        "Sq Ft per Type", "Total Sq Ft",
        "Perimeter Ft", "Total Perimeter Ft"
    ]
    input_values = [
        system_input, elevation_type, total_count,
        bays_wide, bays_tall, opening_width, opening_height,
        sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft
    ]

    for idx, (header, value) in enumerate(zip(inputs_headers, input_values), 1):
        ws[f"A{idx}"] = header
        ws[f"B{idx}"] = value

    # Separate items into groups
    profiles = []
    accessories = []
    manual_outputs = []

    for item in calculated_outputs:
        part_num = item.get('part_number')
        item_type = item.get('type', '').lower()

        if part_num == "N/A" or not part_num:
            if item_type == "profiles":
                profiles.append(item)
            elif item_type == "accessories":
                accessories.append(item)
            else:
                manual_outputs.append(item)
        else:
            if part_num in PART_NUMBER_MAP.get("profiles", []):
                profiles.append(item)
            elif part_num in PART_NUMBER_MAP.get("accessories", []):
                accessories.append(item)
            else:
                manual_outputs.append(item)

    section_col_start = 5  # E
    output_start_row = 1
    grand_total = 0.0

    def write_section(title, items, col_start, row_start, type_label):
        nonlocal grand_total

        ws[f"{get_column_letter(col_start)}{row_start}"] = title
        ws[f"{get_column_letter(col_start)}{row_start + 1}"] = "Part Number"
        ws[f"{get_column_letter(col_start)}{row_start + 2}"] = "Quantity"
        ws[f"{get_column_letter(col_start)}{row_start + 3}"] = "Price"

        output_col_idx = col_start + 1

        for item in items:
            col_letter = get_column_letter(output_col_idx)
            qty = item['quantity']

            if item.get('part_number') and item.get('part_number') != "N/A":
                unit_price, unit_type = get_price_by_part(item['part_number'])
                unit_price = unit_price or 0.0
                unit_type = unit_type or "pcs"
            else:
                unit_price = item.get('price', 0.0)
                unit_type = "units"

            if type_label == "profiles":
                unit_price *= multiplier

            total_cost = qty * unit_price
            grand_total += total_cost

            ws[f"{col_letter}{row_start}"] = item['description']
            ws[f"{col_letter}{row_start + 1}"] = item.get('part_number', 'N/A')
            ws[f"{col_letter}{row_start + 2}"] = f"{qty} {unit_type}"
            ws[f"{col_letter}{row_start + 3}"] = f"${total_cost:.2f}"

            output_col_idx += 1

    write_section("PROFILES", profiles, section_col_start, output_start_row, "profiles")

    accessories_row_start = output_start_row + 6
    write_section("ACCESSORIES", accessories, section_col_start, accessories_row_start, "accessories")

    manual_grouped = {}
    for item in manual_outputs:
        type_key = item.get('type', 'MANUAL').upper()
        manual_grouped.setdefault(type_key, []).append(item)

    current_row = accessories_row_start + 6
    for type_title, items in manual_grouped.items():
        write_section(type_title, items, section_col_start, current_row, "manual")
        current_row += 6

    ws[f"{get_column_letter(section_col_start)}{current_row}"] = "GRAND TOTAL"
    ws[f"{get_column_letter(section_col_start)}{current_row + 1}"] = f"${grand_total:.2f}"

    # --------- AUTO-FIT ALL COLUMNS -----------
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
    # -------------------------------------------

    wb.save("output.xlsx")

    if completion_callback:
        completion_callback("Generated with column widths auto-fitted!", "green")
