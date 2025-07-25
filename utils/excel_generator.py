import os
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, numbers
from openpyxl.utils import get_column_letter
from utils.pricing import get_price_by_part
from data.part_number import PART_NUMBER_MAP

# --- Placeholder for external dependencies ---
# These would typically come from 'data.part_number' and 'utils.pricing'
# For the code to be runnable independently, we'll include mock implementations.

output_file = "output.xlsx"

# --- Helper Functions ---

def _find_row_by_value(ws, column, value, start_row=1, end_row=None, reverse=False):
    """
    Helper function to find the first row number where a cell in a specific column
    contains the given value.
    Args:
        ws (Worksheet): The openpyxl worksheet object.
        column (int): The column index (1-based).
        value (str): The value to search for.
        start_row (int): The row to start searching from (inclusive).
        end_row (int): The row to end searching at (inclusive). Defaults to ws.max_row.
        reverse (bool): If True, search from end_row to start_row.
    Returns:
        int or None: The row number if found, otherwise None.
    """
    end_row = end_row if end_row is not None else ws.max_row
    row_range = range(start_row, end_row + 1)
    if reverse:
        row_range = range(end_row, start_row - 1, -1)

    for r in row_range:
        cell_value = ws.cell(row=r, column=column).value
        if cell_value and str(cell_value).strip() == str(value).strip():
            return r
    return None

def _autofit_columns(ws, start_col, end_col, start_row=1, end_row=None):
    """
    Autofits columns within a specified range based on content length.
    Args:
        ws (Worksheet): The openpyxl worksheet object.
        start_col (int): The starting column index (1-based).
        end_col (int): The ending column index (1-based).
        start_row (int): The row to start measuring from.
        end_row (int): The row to end measuring at. Defaults to ws.max_row.
    """
    end_row = end_row if end_row is not None else ws.max_row
    for col_idx in range(start_col, end_col + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for r in range(start_row, end_row + 1):
            cell_value = ws.cell(row=r, column=col_idx).value
            if cell_value is not None:
                max_len = max(max_len, len(str(cell_value)))
        # Add a small buffer for better readability
        ws.column_dimensions[col_letter].width = max_len + 2

def _clean_trailing_blank_rows(ws, start_row):
    """
    Deletes blank rows from the worksheet starting from a given row
    until a row with content is found or the end of the sheet is reached.
    Args:
        ws (Worksheet): The openpyxl worksheet object.
        start_row (int): The row to start checking for blank rows.
    """
    rows_deleted = 0
    current_row = start_row
    while current_row <= ws.max_row:
        # Check if the entire row is empty
        if all(ws.cell(row=current_row, column=c).value is None for c in range(1, ws.max_column + 1)):
            ws.delete_rows(current_row, 1)
            rows_deleted += 1
            # Do not increment current_row, as rows below shift up
        else:
            current_row += 1 # Move to the next row if current one has content
    if rows_deleted > 0:
        print(f"🧹 Cleaned {rows_deleted} trailing blank rows starting from row {start_row}.")

def _delete_elevation_block(ws, elevation_name, colA, price_col):
    """
    Deletes a specific elevation data block from the worksheet.
    An elevation block starts with "Elevation Type" and ends after its "SYSTEM TOTAL".
    Args:
        ws (Worksheet): The openpyxl worksheet object.
        elevation_name (str): The name of the elevation to delete.
        colA (int): The column index for main headers (e.g., "Elevation Type").
        price_col (int): The column index where "SYSTEM TOTAL" is found.
    """
    delete_start_row = None
    # Find the row where "Elevation Type" and its name appear
    for r in range(1, ws.max_row + 1):
        if (ws.cell(row=r, column=colA).value == "Elevation Type" and
                ws.cell(row=r, column=colA + 1).value == elevation_name):
            delete_start_row = r - 1 # Start deletion from the row above "Elevation Type"
            break

    if delete_start_row is None:
        print(f"ℹ️ Elevation '{elevation_name}' not found for deletion.")
        return

    delete_end_row = None
    # Find the row of "SYSTEM TOTAL" for this elevation
    for r in range(delete_start_row + 1, ws.max_row + 1):
        if ws.cell(row=r, column=price_col).value == "SYSTEM TOTAL":
            delete_end_row = r + 1 # End deletion at the row below "SYSTEM TOTAL"
            break
    
    # Fallback if SYSTEM TOTAL isn't found (e.g., incomplete block)
    if delete_end_row is None:
        # Estimate end if SYSTEM TOTAL isn't found (e.g., 10 rows after start)
        delete_end_row = delete_start_row + 10 
        print(f"⚠️ Could not find 'SYSTEM TOTAL' for '{elevation_name}'. Deleting an estimated block.")

    # Ensure we don't try to delete beyond max_row
    delete_end_row = min(delete_end_row, ws.max_row)

    # Perform deletion
    if delete_end_row >= delete_start_row:
        ws.delete_rows(delete_start_row, delete_end_row - delete_start_row + 1)
        print(f"🗑️ Elevation '{elevation_name}' block deleted from report.")
        # Clean up any blank rows created by the deletion
        # This local cleanup is important, but a final global sweep is also performed
        _clean_trailing_blank_rows(ws, delete_start_row)
    else:
        print(f"⚠️ Deletion range for '{elevation_name}' is invalid ({delete_start_row}-{delete_end_row}). No rows deleted.")


def _recalculate_running_grand_total(ws, price_col):
    """
    Recalculates and updates the "RUNNING GRAND TOTAL" in the worksheet.
    It first removes any existing grand total, then sums all "SYSTEM TOTAL" values,
    and finally writes the new grand total.
    Args:
        ws (Worksheet): The openpyxl worksheet object.
        price_col (int): The column index where "SYSTEM TOTAL" and "RUNNING GRAND TOTAL" are found.
    """
    # Remove any existing RUNNING GRAND TOTAL block
    # Search from bottom up to find the last instance
    for r in range(ws.max_row, 0, -1):
        if ws.cell(row=r, column=price_col).value == "RUNNING GRAND TOTAL":
            ws.delete_rows(r, 2) # Delete the "RUNNING GRAND TOTAL" header and its value
            break # Assuming only one grand total at the end

    running_grand_total = 0.0
    last_system_total_row = None

    # Iterate through all rows to sum up "SYSTEM TOTAL" values
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=price_col).value == "SYSTEM TOTAL":
            last_system_total_row = r
            total_value_cell = ws.cell(row=r + 1, column=price_col).value
            if isinstance(total_value_cell, (float, int)):
                running_grand_total += total_value_cell
            elif isinstance(total_value_cell, str) and total_value_cell.startswith("$"):
                try:
                    running_grand_total += float(total_value_cell.strip("$"))
                except ValueError:
                    print(f"⚠️ Could not parse SYSTEM TOTAL value: {total_value_cell}")

    # Determine the row to write the new RUNNING GRAND TOTAL
    # It should be two rows below the last SYSTEM TOTAL, or at the end of the sheet
    new_gt_row = (last_system_total_row + 3) if last_system_total_row else (ws.max_row + 1)

    ws.cell(row=new_gt_row, column=price_col, value="RUNNING GRAND TOTAL").font = Font(bold=True)
    gt_cell = ws.cell(row=new_gt_row + 1, column=price_col, value=running_grand_total)
    gt_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    print("📈 Running Grand Total recalculated and updated.")


def _write_output_section(ws, title, items, colE, multiplier, system_total_ref, start_output_row):
    """
    Writes a section of calculated outputs (e.g., PROFILES, ACCESSORIES) to the worksheet.
    Args:
        ws (Worksheet): The openpyxl worksheet object.
        title (str): The title of the section (e.g., "PROFILES").
        items (list): A list of dictionaries, each representing an output item.
        colE (int): The starting column index for output details.
        multiplier (float): Price multiplier for profiles.
        system_total_ref (list): A mutable list containing the current system_total.
                                 Used to update the total across calls.
        start_output_row (int): The row where this section should begin writing.
    Returns:
        int: The next available row after this section.
    """
    if not items:
        # If no items, this section doesn't take up space, so the next section
        # should attempt to start at the same row.
        return start_output_row

    current_row = start_output_row

    # Write section title
    ws.cell(row=current_row, column=colE, value=title).font = Font(bold=True)
    
    # Write headers for the section
    headers = ["Description", "Part Number", "Quantity", "Price"]
    for i, header_txt in enumerate(headers):
        ws.cell(row=current_row + 1, column=colE + i, value=header_txt).font = Font(bold=True)
    
    current_row += 2 # Move to the row where item data begins

    # Write each item's details
    for item in items:
        qty = item.get('quantity', 0)
        pn = item.get('part_number')
        
        # Determine unit price and type
        if pn and pn != "N/A":
            unit_price, unit_type = get_price_by_part(pn)
            unit_price = unit_price or 0.0
            unit_type = unit_type or "pcs"
        else: # Handle manual entries without a part number
            unit_price = item.get('price', 0.0)
            unit_type = item.get('unit', 'pcs')

        # Apply multiplier for profiles
        if title == "PROFILES":
            unit_price *= multiplier

        total_item_price = qty * unit_price
        system_total_ref[0] += total_item_price # Update the system total

        ws.cell(row=current_row, column=colE, value=item.get('description', ''))
        ws.cell(row=current_row, column=colE + 1, value=pn or 'N/A')
        ws.cell(row=current_row, column=colE + 2, value=f"{qty} {unit_type}")
        price_cell = ws.cell(row=current_row, column=colE + 3, value=total_item_price)
        price_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        current_row += 1
    
    return current_row + 1 # Return the next available row after this section, plus a blank line


def _delete_summary_section(ws):
    """
    Deletes the existing summary section from the worksheet.
    The summary section is identified by the header "Part Number / Description".
    Args:
        ws (Worksheet): The openpyxl worksheet object.
    """
    summary_header_value = "Part Number / Description"
    summary_start_row = _find_row_by_value(ws, 1, summary_header_value)
    
    if summary_start_row:
        # Find the end of the summary section (first blank row in column A after the header)
        current_row_to_delete = summary_start_row
        while current_row_to_delete <= ws.max_row:
            if ws.cell(row=current_row_to_delete, column=1).value is None:
                break
            current_row_to_delete += 1
        
        # Delete all rows from summary_start_row up to (but not including) the blank row
        if current_row_to_delete > summary_start_row:
            ws.delete_rows(summary_start_row, current_row_to_delete - summary_start_row)
            print("🗑️ Existing summary section cleared.")
            # Clean up any blank rows created by the deletion
            _clean_trailing_blank_rows(ws, summary_start_row)
        else:
            print("ℹ️ Summary header found but no data rows to delete.")
    else:
        print("ℹ️ No existing summary section found to delete.")


# --- Main Functions ---

def create_summary_sheet(excel_path=output_file, json_path='saved_elevations.json'):
    """
    Reads saved elevation data, aggregates quantities and prices by part number (or description for manual),
    and writes a clean summary section into the Excel file.
    """
    try:
        with open(json_path, 'r') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"⚠️ JSON file '{json_path}' not found. Skipping summary sheet creation.")
        return
    except json.JSONDecodeError:
        print(f"⚠️ Error decoding JSON from '{json_path}'. Skipping summary sheet creation.")
        return

    try:
        wb = load_workbook(excel_path)
    except FileNotFoundError:
        print(f"⚠️ Excel file '{excel_path}' not found. Cannot update summary sheet.")
        return

    ws = wb.active
    
    # Always delete the old summary before creating a new one
    _delete_summary_section(ws)

    # If only one or zero elevations, no summary is needed
    if len(data) <= 1:
        wb.save(excel_path)
        print("ℹ️ Only one or zero elevations found. No summary sheet created/updated.")
        return

    # Aggregate outputs from all elevations
    aggregated_summary = {}
    for elev_data in data.values():
        for output in elev_data.get('calculated_outputs', []):
            part_number = output.get('part_number')
            description = output.get('description', '').strip()
            quantity = output.get('quantity', 0)
            price_per_unit_manual = output.get('price') # Price only for manual items

            # Use description as key if no part number or N/A, otherwise use part number
            is_manual_entry = not part_number or part_number == "N/A"
            key = description if is_manual_entry else part_number

            if key not in aggregated_summary:
                aggregated_summary[key] = {
                    'quantity': 0,
                    'price_manual': price_per_unit_manual, # Store manual price if applicable
                    'is_manual': is_manual_entry,
                    'description': description # Keep description for non-part-number items
                }
            aggregated_summary[key]['quantity'] += quantity

    # Prepare rows for the summary table
    summary_rows_to_write = []
    for key, entry in aggregated_summary.items():
        qty = entry['quantity']
        if entry['is_manual']:
            price_per_unit = entry['price_manual'] or 0.0
        else:
            price_per_unit, _ = get_price_by_part(key) # Key is the part number here
            price_per_unit = price_per_unit or 0.0

        total_price = price_per_unit * qty
        summary_rows_to_write.append((key, qty, total_price))

    # Determine where to insert the summary (after the last RUNNING GRAND TOTAL)
    # Assuming RUNNING GRAND TOTAL is in column E+3 (which is column 8)
    last_grand_total_row = _find_row_by_value(ws, 8, "RUNNING GRAND TOTAL", reverse=True)
    # Start 3 rows after the grand total, or 2 rows after the current max row if no grand total
    start_row_for_summary = (last_grand_total_row + 3) if last_grand_total_row else (ws.max_row + 2)

    # Write headers for the summary table
    ws.cell(row=start_row_for_summary, column=1, value="Part Number / Description").font = Font(bold=True)
    ws.cell(row=start_row_for_summary, column=2, value="Total Quantity").font = Font(bold=True)
    ws.cell(row=start_row_for_summary, column=3, value="Total Price").font = Font(bold=True)

    # Write summary data rows
    for idx, (item_key, qty, total) in enumerate(summary_rows_to_write, start=start_row_for_summary + 1):
        ws.cell(row=idx, column=1, value=item_key)
        ws.cell(row=idx, column=2, value=qty)
        price_cell = ws.cell(row=idx, column=3, value=total)
        price_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    # Autofit columns for the summary section
    _autofit_columns(ws, 1, 3, start_row_for_summary, start_row_for_summary + len(summary_rows_to_write))

    # Perform a final cleanup of any blank rows from the top before saving
    _clean_trailing_blank_rows(ws, 1)
    wb.save(excel_path)
    print(f"✅ Summary sheet updated in {excel_path}.")


def generate_excel_report(
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, mode=None, reset=False, all_elevations=None,
    delete_elevation_type=None
):
    """
    Generates or updates an Excel report with detailed elevation inputs and calculated outputs.
    Manages adding new elevation blocks, deleting existing ones, and recalculating totals.
    """
    # Define column constants for clarity
    COL_A, COL_B, COL_E = 1, 2, 5
    PRICE_COL = COL_E + 3 # Column for prices (E + 3 = H)

    # Determine price multiplier based on finish input
    multiplier = {"clear": 1.0, "black": 1.1, "paint": 1.2}.get(finish_input.lower(), 1.0)

    # Load or create the workbook
    if reset or mode == "new" or not os.path.exists(output_file):
        wb = Workbook()
        ws = wb.active
        ws.title = "Report"
        print(f"📄 Created new Excel workbook: {output_file}")
    else:
        wb = load_workbook(output_file)
        ws = wb.active
        print(f"📂 Loaded existing Excel workbook: {output_file}")

    # --- Handle Deletion Mode ---
    if delete_elevation_type:
        # Delete summary first to ensure grand total calculation is clean
        _delete_summary_section(ws)
        _delete_elevation_block(ws, delete_elevation_type, COL_A, PRICE_COL)
        _recalculate_running_grand_total(ws, PRICE_COL)
        # Perform a final cleanup of any blank rows from the top before saving
        _clean_trailing_blank_rows(ws, 1)
        wb.save(output_file)
        create_summary_sheet(excel_path=output_file) # Update summary after deletion
        if completion_callback:
            completion_callback()
        return

    # --- Handle Adding/Updating Elevations ---
    # Delete the summary section first to ensure it's recreated at the end based on fresh data
    _delete_summary_section(ws)

    # If not in reset mode, and an elevation type is provided, delete its existing block
    # to prevent duplicates when updating an existing elevation.
    if not reset and elevation_type:
        _delete_elevation_block(ws, elevation_type, COL_A, PRICE_COL)
    
    # Determine the starting row for the new elevation block
    start_row_for_new_block = 1
    if not reset:
        # Check if the sheet is effectively empty after deletions
        is_sheet_empty = True
        for r in range(1, ws.max_row + 1):
            if any(ws.cell(row=r, column=c).value is not None for c in range(1, ws.max_column + 1)):
                is_sheet_empty = False
                break

        if is_sheet_empty:
            start_row_for_new_block = 1
        else:
            last_grand_total_row = _find_row_by_value(ws, PRICE_COL, "RUNNING GRAND TOTAL", reverse=True)
            # If a grand total exists, start 3 rows after it. Otherwise, append to the end.
            start_row_for_new_block = (last_grand_total_row + 3) if last_grand_total_row else (ws.max_row + 2)
    else: # If reset is True, clear the entire sheet and start at row 1
        ws.delete_rows(1, ws.max_row)
        print("🧹 Worksheet cleared for new report.")
        start_row_for_new_block = 1

    # Write input headers and values (left side of the report)
    input_data = [
        ("System Input", system_input), ("Elevation Type", elevation_type), ("Total Count", total_count),
        ("# Bays Wide", bays_wide), ("# Bays Tall", bays_tall), ("Opening Width", opening_width),
        ("Opening Height", opening_height), ("Sq Ft per Type", sqft_per_type), ("Total Sq Ft", total_sqft),
        ("Perimeter Ft", perimeter_ft), ("Total Perimeter Ft", total_perimeter_ft)
    ]
    for i, (header, value) in enumerate(input_data):
        ws.cell(row=start_row_for_new_block + i, column=COL_A, value=header).font = Font(bold=True)
        ws.cell(row=start_row_for_new_block + i, column=COL_B, value=value)

    # Initialize system total for the current elevation block (passed by reference to helper)
    current_system_total = [0.0]

    # Process and write calculated outputs (right side of the report)
    if all_elevations and not reset: # Only write outputs if not resetting and there's data
        # Categorize outputs based on part number or type
        profiles = []
        accessories = []
        other_manual_outputs = [] # For items without specific categories or N/A part numbers

        for item in calculated_outputs:
            part_number = item.get('part_number')
            item_type = item.get('type', '').lower()

            if part_number and part_number != "N/A":
                if part_number in PART_NUMBER_MAP.get("profiles", []):
                    profiles.append(item)
                elif part_number in PART_NUMBER_MAP.get("accessories", []):
                    accessories.append(item)
                else:
                    # If it has a part number but isn't in known categories, treat as other
                    other_manual_outputs.append(item)
            else:
                # If no part number or N/A, it's a manual entry. Categorize by its 'type'.
                other_manual_outputs.append(item)

        # Start output sections at the same row as input headers
        output_section_current_row = start_row_for_new_block

        # Write each section, updating the current row for the next section
        output_section_current_row = _write_output_section(ws, "PROFILES", profiles, COL_E, multiplier, current_system_total, output_section_current_row)
        output_section_current_row = _write_output_section(ws, "ACCESSORIES", accessories, COL_E, multiplier, current_system_total, output_section_current_row)

        # Group and write other manual outputs by their 'type'
        grouped_other = {}
        for item in other_manual_outputs:
            # Use the 'type' field from the item, default to 'MANUAL ITEMS'
            group_key = item.get('type', 'MANUAL ITEMS').upper()
            grouped_other.setdefault(group_key, []).append(item)
        
        for grp_title, grp_items in grouped_other.items():
            # For these, the multiplier is not applied, so pass 1.0
            output_section_current_row = _write_output_section(ws, grp_title, grp_items, COL_E, 1.0, current_system_total, output_section_current_row)

    # Write the SYSTEM TOTAL for the current elevation block
    # Determine the row for SYSTEM TOTAL (it should be after all sections for this elevation)
    system_total_row = ws.max_row + 2
    ws.cell(row=system_total_row, column=PRICE_COL, value="SYSTEM TOTAL").font = Font(bold=True)
    sys_total_cell = ws.cell(row=system_total_row + 1, column=PRICE_COL, value=current_system_total[0])
    sys_total_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    print(f"📊 System Total for '{elevation_type}': ${current_system_total[0]:.2f}")

    # Recalculate and update the overall RUNNING GRAND TOTAL
    _recalculate_running_grand_total(ws, PRICE_COL)

    # Autofit all columns that might contain data
    _autofit_columns(ws, COL_A, PRICE_COL, 1, ws.max_row)

    # Perform a final cleanup of any blank rows from the top before saving
    _clean_trailing_blank_rows(ws, 1)
    wb.save(output_file)
    print(f"✅ Excel report '{output_file}' updated successfully.")

    # Update the summary sheet after the main report is generated/updated
    create_summary_sheet(excel_path=output_file)

    if completion_callback:
        completion_callback()
