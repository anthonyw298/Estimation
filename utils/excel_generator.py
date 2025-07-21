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

    # OUTPUT labels in D1 to D4
    ws["D1"] = "OUTPUT"
    ws["D2"] = "Part Number"
    ws["D3"] = "Quantity"
    ws["D4"] = "Price"

    # Calculate total costs
    total_costs = []
    unit_types = []  # keep unit type for each item

    for item in calculated_outputs:
        part_num = item.get('part_number') or PART_NUMBER_MAP.get(item['description'], "")
        qty = item['quantity']
        unit_price, unit_type = get_price_by_part(part_num)
        unit_price = unit_price or 0.0
        unit_type = unit_type or "pcs"
        unit_types.append(unit_type)

        if item.get('type') == 'profiles':
            unit_price *= multiplier

        total_cost = qty * unit_price
        total_costs.append(total_cost)

    # Write output columns starting at E1, E2, ...
    output_col_idx = 5  # E
    output_start_row = 1

    for idx, item in enumerate(calculated_outputs):
        col_letter = get_column_letter(output_col_idx)

        part_num = item.get('part_number') or PART_NUMBER_MAP.get(item['description'], "")
        qty = item['quantity']
        unit_type = unit_types[idx]

        # Description at top
        ws[f"{col_letter}{output_start_row}"] = item['description']
        # Part number under OUTPUT labels
        ws[f"{col_letter}{output_start_row + 1}"] = part_num
        # Quantity with unit label
        ws[f"{col_letter}{output_start_row + 2}"] = f"{qty} {unit_type}"
        # Price formatted
        ws[f"{col_letter}{output_start_row + 3}"] = f"${total_costs[idx]:.2f}"

        output_col_idx += 1

    # GRAND TOTAL: two rows up from previous
    grand_total_col_letter = get_column_letter(output_col_idx)
    grand_total_row = output_start_row + 2  # moved up by 2 more rows
    ws[f"{grand_total_col_letter}{grand_total_row}"] = "GRAND TOTAL"
    ws[f"{grand_total_col_letter}{grand_total_row + 1}"] = f"${sum(total_costs):.2f}"

    # Make column C wide for a visual gap
    ws.column_dimensions['C'].width = 15

    # Auto-size other columns
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'C':
            continue  # skip gap column
        max_length = max((len(str(cell.value)) if cell.value else 0) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_length + 2, 10)

    wb.save("output.xlsx")
    if completion_callback:
        completion_callback("Excel file 'output.xlsx' generated successfully!", "green")
