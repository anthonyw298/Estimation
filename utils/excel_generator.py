import openpyxl
from openpyxl.utils import get_column_letter
from data.part_number import PART_NUMBER_MAP
from utils.pricing import get_price_by_part

def generate_excel_report(
    system_input: str,
    finish_input: str,  # new finish input
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
    calculated_outputs: list,  # list of dicts with 'description', 'quantity', 'part_number' keys, may contain 'profile_type'
    completion_callback=None  # Optional callback for status updates
):
    if system_input != "YES 45TU FRONT SET(OG)":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = f"System '{system_input}' not matched. Empty file created."
        try:
            wb.save("output.xlsx")
            if completion_callback:
                completion_callback(f"System not matched. Empty 'output.xlsx' created.", "orange")
        except Exception as e:
            if completion_callback:
                completion_callback(f"Error saving empty file: {e}", "red")
        return

    finish_multiplier_map = {
        "clear": 1.0,
        "black": 1.1,
        "paint": 1.2
    }
    multiplier = finish_multiplier_map.get(finish_input.lower(), 1.0)

    try:
        wb = openpyxl.Workbook()
        ws = wb.active

        # Define headers
        inputs_headers = [
            "System Input", "Elevation Type", "Total Count",
            "# Bays Wide", "# Bays Tall",
            "Opening Width", "Opening Height",
            "Sq Ft per Type", "Total Sq Ft", "Perimeter Ft", "Total Perimeter Ft"
        ]
        outputs_headers = [item['description'] for item in calculated_outputs]
        headers = inputs_headers + outputs_headers

        # Write headers row 1 (inputs + output descriptions)
        for idx, header in enumerate(headers, 1):
            ws[f"{get_column_letter(idx)}1"] = header

        # Write part numbers in row 2 (blank for input columns)
        for idx in range(1, len(inputs_headers) + 1):
            ws[f"{get_column_letter(idx)}2"] = ""  # blank for inputs columns

        for col_idx, item in enumerate(calculated_outputs, start=len(inputs_headers) + 1):
            part_num = item.get('part_number') or PART_NUMBER_MAP.get(item['description'], "")
            ws[f"{get_column_letter(col_idx)}2"] = part_num

        # Write input values + output quantities in row 3
        input_values = [
            system_input, elevation_type, total_count,
            bays_wide, bays_tall, opening_width, opening_height,
            sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft
        ]
        output_values = [item['quantity'] for item in calculated_outputs]
        all_values = input_values + output_values

        for idx, val in enumerate(all_values, 1):
            ws[f"{get_column_letter(idx)}3"] = val

        # Calculate prices for outputs with finish multiplier applied only if profile_type exists
        prices = []
        for item in calculated_outputs:
            part_num = item.get('part_number') or PART_NUMBER_MAP.get(item['description'])
            if part_num:
                price = get_price_by_part(part_num)
                print(price)
                if price is None:
                    price = 0.0

                if item.get('type') == 'profiles':
                    price *= multiplier
                prices.append(price)
            else:
                prices.append(0.0)

        # Calculate total costs per part
        total_costs = []
        for qty, price in zip(output_values, prices):
            try:
                total_costs.append(float(qty) * float(price))
            except Exception:
                total_costs.append(0.0)

        # Write total cost row 4 (blank inputs columns)
        for idx in range(1, len(inputs_headers) + 1):
            ws[f"{get_column_letter(idx)}4"] = ""  # blank for inputs columns

        for idx, cost in enumerate(total_costs, start=len(inputs_headers) + 1):
            ws[f"{get_column_letter(idx)}4"] = cost

        ws["A4"] = "Total Cost ($)"


        # Auto-size columns
        for idx, header in enumerate(headers, 1):
            col_letter = get_column_letter(idx)
            max_len = len(str(header))
            for row in range(1, 5):
                cell_val = ws[f"{col_letter}{row}"].value
                if cell_val is not None:
                    max_len = max(max_len, len(str(cell_val)))
            ws.column_dimensions[col_letter].width = max_len + 2

        wb.save("output.xlsx")
        if completion_callback:
            completion_callback("Excel file 'output.xlsx' generated successfully!", "green")
        print("Excel file saved as 'output.xlsx'")

    except Exception as e:
        if completion_callback:
            completion_callback(f"An error occurred during Excel generation: {e}", "red")
        print(f"Error during Excel generation: {e}")
