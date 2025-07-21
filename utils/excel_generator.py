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
    # If system does not match, output an empty file with a message.
    if system_input != "YES 45TU FRONT SET(OG)":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = f"System '{system_input}' not matched. Empty file created."
        wb.save("output.xlsx")
        if completion_callback:
            completion_callback("System not matched. Empty 'output.xlsx' created.", "orange")
        return

    # Apply finish multiplier.
    finish_multiplier_map = {
        "clear": 1.0,
        "black": 1.1,
        "paint": 1.2
    }
    multiplier = finish_multiplier_map.get(finish_input.lower(), 1.0)

    wb = openpyxl.Workbook()
    ws = wb.active

    # Inputs block in columns A & B
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

    # Split items into profiles and accessories
    profiles = []
    accessories = []

    for item in calculated_outputs:
        part_num = item.get('part_number') or PART_NUMBER_MAP.get(item['description'], "")
        if part_num in PART_NUMBER_MAP["profiles"]:
            profiles.append(item)
        else:
            accessories.append(item)

    section_col_start = 5  # column E
    output_start_row = 1

    grand_total = 0.0

    def write_section(title, items, col_start, row_start, type_label):
        nonlocal grand_total

        ws[f"{get_column_letter(col_start)}{row_start}"] = title
        ws[f"{get_column_letter(col_start)}{row_start + 1}"] = "Part Number"
        ws[f"{get_column_letter(col_start)}{row_start + 2}"] = "Quantity"
        ws[f"{get_column_letter(col_start)}{row_start + 3}"] = "Price"

        output_col_idx = col_start + 1  # start writing next to labels

        for idx, item in enumerate(items):
            col_letter = get_column_letter(output_col_idx)

            part_num = item.get('part_number') or PART_NUMBER_MAP.get(item['description'], "")
            qty = item['quantity']
            unit_price, unit_type = get_price_by_part(part_num)
            unit_price = unit_price or 0.0
            unit_type = unit_type or "pcs"

            if type_label == "profiles":
                unit_price *= multiplier

            total_cost = qty * unit_price
            grand_total += total_cost

            # Description at top
            ws[f"{col_letter}{row_start}"] = item['description']
            # Part number under labels
            ws[f"{col_letter}{row_start + 1}"] = part_num
            # Quantity with unit
            ws[f"{col_letter}{row_start + 2}"] = f"{qty} {unit_type}"
            # Price
            ws[f"{col_letter}{row_start + 3}"] = f"${total_cost:.2f}"

            output_col_idx += 1

    # Write profiles section
    write_section("PROFILES", profiles, section_col_start, output_start_row, "profiles")

    # Write accessories section, shifted down by 6 rows
    accessories_row_start = output_start_row + 6
    write_section("ACCESSORIES", accessories, section_col_start, accessories_row_start, "accessories")

    # GRAND TOTAL
    grand_total_col = get_column_letter(section_col_start)
    ws[f"{grand_total_col}{accessories_row_start + 6}"] = "GRAND TOTAL"
    ws[f"{grand_total_col}{accessories_row_start + 7}"] = f"${grand_total:.2f}"

    # Make column C wide for a gap
    ws.column_dimensions['C'].width = 15

    # Auto-size all columns except C
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'C':
            continue
        max_length = max((len(str(cell.value)) if cell.value else 0) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_length + 2, 10)

    wb.save("output.xlsx")
    if completion_callback:
        completion_callback("Excel file 'output.xlsx' generated with profiles and accessories!", "green")
