# utils/excel_generator.py

import openpyxl
from openpyxl.utils import get_column_letter

def generate_excel_report(
    system_input: str,
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
    calculated_outputs: dict, # This is the key: it expects a dictionary of outputs
    completion_callback # A callback function to update UI status
):
    """
    Generates an Excel file based on provided inputs and pre-calculated outputs.
    Updates the UI via a callback function.
    """
    # --- System Input Validation (Simplified as main.py handles initial check) ---
    if system_input != "YES 45TU Front Set(OG)":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = f"System '{system_input}' not matched. Empty file created."
        try:
            wb.save("output.xlsx")
            completion_callback(f"System not matched. Empty 'output.xlsx' created.", "orange")
        except Exception as e:
            completion_callback(f"Error saving empty file: {e}", "red")
        return # Exit the function if system doesn't match

    # --- Proceed with Excel generation if system matches ---
    try:
        wb = openpyxl.Workbook()
        ws = wb.active

        # Define input headers
        inputs_headers = [
            "System Input", "Elevation Type", "Total Count",
            "# Bays Wide", "# Bays Tall",
            "Opening Width", "Opening Height",
            "Sq Ft per Type", "Total Sq Ft", "Perimeter Ft", "Total Perimeter Ft"
        ]

        # Get output headers from the keys of the calculated_outputs dictionary
        # This makes it flexible for other systems if they have different outputs
        outputs_headers = list(calculated_outputs.keys())

        # Combine input and output headers and write them to the first row
        headers = inputs_headers + outputs_headers
        for idx, header in enumerate(headers):
            ws[f"{get_column_letter(idx + 1)}1"] = header

        # Prepare input values for row 2
        input_values = [
            system_input, elevation_type, total_count,
            bays_wide, bays_tall, opening_width, opening_height,
            sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft
        ]

        # Get output values in the correct order based on outputs_headers
        output_values = [calculated_outputs[key] for key in outputs_headers]

        # Combine inputs and outputs and write them to row 2
        all_values = input_values + output_values
        for idx, val in enumerate(all_values):
            ws[f"{get_column_letter(idx + 1)}2"] = val

        # Auto-size each column based on max length of header or data
        for idx, header in enumerate(headers, 1):
            col_letter = get_column_letter(idx)
            cell_value = str(ws[f"{col_letter}2"].value) if ws[f"{col_letter}2"].value is not None else ""
            ws.column_dimensions[col_letter].width = max(len(header), len(cell_value)) + 2

        # Save workbook
        wb.save("output.xlsx")
        completion_callback("Excel file 'output.xlsx' generated successfully!", "green")
        print("Excel file saved as 'output.xlsx'") # Also print to console for confirmation

    except Exception as e:
        completion_callback(f"An error occurred during Excel generation: {e}", "red")
        print(f"Error during Excel generation: {e}")