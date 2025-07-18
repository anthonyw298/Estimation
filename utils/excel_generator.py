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

    # Inputs in A & B
    inputs_headers = [
        "System Input", "Elevation Type", "Total Count",
        "# Bays Wide", "# Bays Tall",
        "Opening Width", "Opening Height",
        "Sq Ft per Type", "Total Sq Ft", "Perimeter Ft", "Total Perimeter Ft"
    ]
    input_values = [
        system_input, elevation_type, total_count,
        bays_wide, bays_tall, opening_width, opening_height,
        sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft
    ]

    for idx, (header, value) in enumerate(zip(inputs_headers, input_values), 1):
        ws[f"A{idx}"] = header
        ws[f"B{idx}"] = value

    # OUTPUT label in D1, labels under it
    ws["D1"] = "OUTPUT"
    ws["D2"] = "Part Number"
    ws["D3"] = "Quantity"
    ws["D4"] = "Price"

    # Compute costs
    total_costs = []
    for item in calculated_outputs:
        part_num = item.get('part_number') or PART_NUMBER_MAP.get(item['description'])
        qty = item['quantity']
        price = get_price_by_part(part_num) or 0.0
        if item.get('type') == 'profiles':
            price *= multiplier
        total_cost = qty * price
        total_costs.append(total_cost)

    # Outputs side-by-side starting at column E (no skipping)
    output_col_idx = 5  # E
    output_start_row = 1

    for idx, item in enumerate(calculated_outputs):
        col_letter = get_column_letter(output_col_idx)

        # Description at top
        ws[f"{col_letter}{output_start_row}"] = item['description']

        # Part number under OUTPUT labels
        part_num = item.get('part_number') or PART_NUMBER_MAP.get(item['description'], "")
        ws[f"{col_letter}{output_start_row + 1}"] = part_num

        # Quantity with unit
        qty = item['quantity']
        qty_unit = "ft" if isinstance(qty, float) and not qty.is_integer() else "pcs"
        ws[f"{col_letter}{output_start_row + 2}"] = f"{qty} {qty_unit}"

        # Cost formatted
        ws[f"{col_letter}{output_start_row + 3}"] = f"${total_costs[idx]:.2f}"

        output_col_idx += 1  # move to next column, no skipping

    # Move GRAND TOTAL two more rows UP (4 rows up total from previous position)
    grand_total_col_letter = get_column_letter(output_col_idx)
    grand_total_row = output_start_row + 2  # Moved two more rows up (from 4 to 2)
    ws[f"{grand_total_col_letter}{grand_total_row}"] = "GRAND TOTAL"
    ws[f"{grand_total_col_letter}{grand_total_row + 1}"] = f"${sum(total_costs):.2f}"

    # Widen blank column C
    ws.column_dimensions['C'].width = 15  # wider gap

    # Auto-size others
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'C':
            continue  # already set
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_length + 2, 10)

    wb.save("output.xlsx")
    if completion_callback:
        completion_callback("Excel file 'output.xlsx' generated successfully!", "green")
