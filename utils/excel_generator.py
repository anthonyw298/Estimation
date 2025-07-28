import os
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, numbers
from openpyxl.utils import get_column_letter
from utils.pricing import get_price_by_part
from data.part_number import PART_NUMBER_MAP

output_file = "output.xlsx"

# --- Helper Functions ---

def _find_row_by_value(ws, column, value, start_row=1, end_row=None, reverse=False):
    """Finds the first row containing a specific value in a given column."""
    end_row = end_row if end_row is not None else ws.max_row
    row_range = range(start_row, end_row + 1)
    if reverse: row_range = range(end_row, start_row - 1, -1)
    for r in row_range:
        cell_value = ws.cell(row=r, column=column).value
        if cell_value and str(cell_value).strip() == str(value).strip(): return r
    return None

def _autofit_columns(ws, start_col, end_col, start_row=1, end_row=None):
    """Autofits columns within a specified range based on content length."""
    end_row = end_row if end_row is not None else ws.max_row
    for col_idx in range(start_col, end_col + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max((len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(start_row, end_row + 1)), default=0)
        ws.column_dimensions[col_letter].width = max_len + 2

def _clean_trailing_blank_rows(ws, start_row):
    """Deletes blank rows from the worksheet starting from a given row."""
    rows_deleted = 0
    current_row = start_row
    while current_row <= ws.max_row:
        if all(ws.cell(row=current_row, column=c).value is None for c in range(1, ws.max_column + 1)):
            ws.delete_rows(current_row, 1); rows_deleted += 1
        else: current_row += 1
    if rows_deleted > 0: print(f"🧹 Cleaned {rows_deleted} trailing blank rows starting from row {start_row}.")

def _delete_elevation_block(ws, elevation_name, colA, price_col):
    """Deletes a specific elevation data block from the worksheet."""
    delete_start_row = _find_row_by_value(ws, colA, "Elevation Type", end_row=ws.max_row)
    if delete_start_row and ws.cell(row=delete_start_row, column=colA + 1).value == elevation_name:
        delete_start_row -= 1
    else: print(f"ℹ️ Elevation '{elevation_name}' not found for deletion."); return

    delete_end_row = (_find_row_by_value(ws, price_col, "SYSTEM TOTAL", start_row=delete_start_row + 1) or delete_start_row + 10) + 1
    delete_end_row = min(delete_end_row, ws.max_row)

    if delete_end_row >= delete_start_row:
        ws.delete_rows(delete_start_row, delete_end_row - delete_start_row + 1)
        print(f"🗑️ Elevation '{elevation_name}' block deleted from report."); _clean_trailing_blank_rows(ws, delete_start_row)
    else: print(f"⚠️ Deletion range for '{elevation_name}' is invalid ({delete_start_row}-{delete_end_row}). No rows deleted.")

def _recalculate_running_grand_total(ws, price_col):
    """Recalculates and updates the 'RUNNING GRAND TOTAL' in the worksheet."""
    for r in range(ws.max_row, 0, -1):
        if ws.cell(row=r, column=price_col).value == "RUNNING GRAND TOTAL":
            ws.delete_rows(r, 2); break

    running_grand_total = 0.0
    last_system_total_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=price_col).value == "SYSTEM TOTAL":
            last_system_total_row = r
            val = ws.cell(row=r + 1, column=price_col).value
            if isinstance(val, (float, int)): running_grand_total += val
            elif isinstance(val, str) and val.startswith("$"):
                try: running_grand_total += float(val.strip("$"))
                except ValueError: print(f"⚠️ Could not parse SYSTEM TOTAL value: {val}")

    new_gt_row = (last_system_total_row + 3) if last_system_total_row else (ws.max_row + 1)
    ws.cell(row=new_gt_row, column=price_col, value="RUNNING GRAND TOTAL").font = Font(bold=True)
    ws.cell(row=new_gt_row + 1, column=price_col, value=running_grand_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    print("📈 Running Grand Total recalculated and updated.")

def _write_output_section(ws, title, items, colE, multiplier, system_total_ref, start_output_row):
    """Writes a section of calculated outputs (e.g., PROFILES, ACCESSORIES) to the worksheet."""
    if not items: return start_output_row

    current_row = start_output_row
    ws.cell(row=current_row, column=colE, value=title).font = Font(bold=True)
    for i, h in enumerate(["Description", "Part Number", "Quantity", "Price"]):
        ws.cell(row=current_row + 1, column=colE + i, value=h).font = Font(bold=True)
    current_row += 2

    for item in items:
        qty = item.get('quantity', 0)
        pn = item.get('part_number')
        manual = item.get('manual', False) # Explicitly get the manual flag

        total_item_price = 0.0
        unit_type = "pcs"

        if manual: # If it's explicitly marked as manual
            if pn and pn != "N/A": # If manual and has a part number, try to get price from part data
                try:
                    # Pass the 'manual' flag to get_price_by_part if its logic needs to differ for manual items
                    price_from_part, unit = get_price_by_part(pn, qty, group=True) 
                    # If get_price_by_part returns a valid price, use it, otherwise fall back to 'price' in item
                    # No longer multiply by qty here, get_price_by_part should return total for manual items
                    total_item_price = (price_from_part if price_from_part is not None else item.get('price', 0.0) * qty)
                    unit_type = unit or item.get('unit', 'pcs')
                except Exception as e:
                    print(f"⚠️ Error getting price for manual item with PN '{pn}': {e}. Using provided price.")
                    total_item_price = item.get('price', 0.0) * qty
                    unit_type = item.get('unit', 'pcs')
            else: # Manual entry without a part number (or "N/A")
                total_item_price = item.get('price', 0.0) * qty
                unit_type = item.get('unit', 'pcs')
        else: # Not a manual entry, always get price from part data
            # get_price_by_part now returns the total cost for the requested qty and unit_type
            total_item_price, unit_type = get_price_by_part(pn, qty)
            total_item_price = total_item_price or 0.0
            unit_type = unit_type or "pcs"
        
        # Apply multiplier for profiles to the total calculated item price
        if title == "PROFILES":
            total_item_price *= multiplier

        system_total_ref[0] += total_item_price

        ws.cell(row=current_row, column=colE, value=item.get('description', ''))
        ws.cell(row=current_row, column=colE + 1, value=pn or 'N/A')
        ws.cell(row=current_row, column=colE + 2, value=f"{qty} {unit_type}")
        ws.cell(row=current_row, column=colE + 3, value=total_item_price).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        current_row += 1
    return current_row + 1

def _delete_summary_section(ws):
    """Deletes the existing summary section from the worksheet."""
    # Only attempt to delete if there's more than one row (i.e., not a completely empty sheet)
    if ws.max_row <= 1: # A new sheet often has 1 row by default, which is empty.
        print("ℹ️ Worksheet is largely empty. Skipping summary section deletion.")
        return

    summary_start_row = _find_row_by_value(ws, 1, "Part Number / Description")
    if summary_start_row:
        current_row_to_delete = summary_start_row
        while current_row_to_delete <= ws.max_row and ws.cell(row=current_row_to_delete, column=1).value is not None:
            current_row_to_delete += 1
        if current_row_to_delete > summary_start_row:
            ws.delete_rows(summary_start_row, current_row_to_delete - summary_start_row)
            print("🗑️ Existing summary section cleared."); _clean_trailing_blank_rows(ws, summary_start_row)
        else: print("ℹ️ Summary header found but no data rows to delete.")
    else: print("ℹ️ No existing summary section found to delete.")

# --- Main Functions ---

def create_summary_sheet(excel_path=output_file, json_path='saved_elevations.json'):
    """Reads elevation data, aggregates quantities and prices by part number (or description for manual),
    and writes a clean summary section into the Excel file."""
    try:
        data = json.load(open(json_path, 'r'))
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"⚠️ Error loading JSON file '{json_path}': {e}. Skipping summary."); return

    wb = None
    try:
        wb = load_workbook(excel_path)
        ws = wb.active
    except Exception as e:
        print(f"⚠️ Excel file '{excel_path}' not found or corrupted for summary: {e}. Cannot update summary sheet.")
        return
    
    # Only attempt to delete summary section if the workbook is not newly created/reset (already has content)
    # This check is less direct here than in generate_excel_report, as create_summary_sheet is called independently.
    # We can rely on _delete_summary_section's internal check.
    _delete_summary_section(ws) 
    
    if len(data) <= 1:
        try:
            if wb: wb.save(excel_path)
            print("ℹ️ Only one or zero elevations found. No summary sheet created/updated.")
        except Exception as save_err:
            print(f"❌ Error saving workbook after summary check: {save_err}")
        return

# === FIRST: Aggregate all quantities by part number or description ===

    aggregated_summary = {}

    for elev_data in data.values():
        for output in elev_data.get('calculated_outputs', []):
            part_number = output.get('part_number', '').strip()
            description = output.get('description', '').strip()
            quantity = output.get('quantity', 0)
            manual = output.get('manual', False)

            # Decide the key: part_number if valid and not manual, else description
            if manual or not part_number or part_number == "N/A":
                key = description
            else:
                key = part_number

            if key not in aggregated_summary:
                aggregated_summary[key] = {
                    'quantity': 0,
                    'description': description,
                    'part_number': part_number,
                    'manual': manual,
                    'price': output.get('price', 0.0)  # Save any manual price for later
                }

            aggregated_summary[key]['quantity'] += quantity

    # === SECOND: For each unique item, run get_price_by_part ===

    for key, item in aggregated_summary.items():
        quantity = item['quantity']
        manual = item['manual']
        part_number = item['part_number']

        if manual:
            # If manual, use saved price or get price by part if a valid part number exists
            if part_number and part_number != "N/A":
                print('hello world')
                try:
                    total_price, _ = get_price_by_part(part_number, quantity, group=True,summary=True)
                    print('helloworld')
                    item['total_cost'] = total_price if total_price is not None else item['price']
                except Exception:
                    print('1234565')
                    item['total_cost'] = item['price'] * quantity
            else:
                item['total_cost'] = item['price'] * quantity
        else:
            # Auto items — always call get_price_by_part
            print(quantity,part_number)
            total_price, _ = get_price_by_part(part_number, quantity,summary=True)
            print(total_price)
            item['total_cost'] = total_price if total_price is not None else 0.0

    # ✅ Now aggregated_summary has:
    # {
    #   key: {
    #     'quantity': total_qty,
    #     'description': ...,
    #     'part_number': ...,
    #     'manual': ...,
    #     'price': ...,
    #     'total_cost': ... <-- Final cost
    #   }
    # }

    # === OPTIONAL: Print or return it ===


    summary_rows_to_write = []
    for key, entry in aggregated_summary.items():
        # The total cost is already aggregated in 'total_cost'
        summary_rows_to_write.append((key, entry['quantity'], entry['total_cost']))

    last_gt_row = _find_row_by_value(ws, 8, "RUNNING GRAND TOTAL", reverse=True)
    start_row = (last_gt_row + 3) if last_gt_row else (ws.max_row + 2)

    ws.cell(row=start_row, column=1, value="Part Number / Description").font = Font(bold=True)
    ws.cell(row=start_row, column=2, value="Total Quantity").font = Font(bold=True)
    ws.cell(row=start_row, column=3, value="Total Price").font = Font(bold=True)

    for idx, (item_key, qty, total) in enumerate(summary_rows_to_write, start=start_row + 1):
        ws.cell(row=idx, column=1, value=item_key)
        ws.cell(row=idx, column=2, value=qty)
        price_cell = ws.cell(row=idx, column=3, value=total)
        price_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    _autofit_columns(ws, 1, 3, start_row, start_row + len(summary_rows_to_write))
    _clean_trailing_blank_rows(ws, 1)
    try:
        print(aggregated_summary)
        wb.save(excel_path); print(f"✅ Summary sheet updated in {excel_path}.")
    except Exception as save_err:
        print(f"❌ Error saving summary sheet to '{excel_path}': {save_err}")

def generate_excel_report(
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, mode=None, reset=False, all_elevations=None,
    delete_elevation_type=None
):
    """Generates or updates an Excel report with detailed elevation inputs and calculated outputs."""
    COL_A, COL_B, COL_E, PRICE_COL = 1, 2, 5, 8
    multiplier = {"clear": 1.0, "black": 1.1, "paint": 1.2}.get(finish_input.lower(), 1.0)

    wb = None
    try:
        is_new_workbook = False # Flag to track if a new workbook is created

        if reset or mode == "new":
            wb = Workbook()
            is_new_workbook = True
            print(f"📄 Created new Excel workbook for reset/new mode: {output_file}")
        elif os.path.exists(output_file):
            try:
                wb = load_workbook(output_file)
                print(f"📂 Loaded existing Excel workbook: {output_file}")
            except Exception as e:
                print(f"⚠️ Error loading existing Excel file '{output_file}': {e}. Creating a new one as fallback.")
                wb = Workbook()
                is_new_workbook = True
                print(f"📄 Created new Excel workbook as fallback: {output_file}")
        else:
            wb = Workbook()
            is_new_workbook = True
            print(f"📄 Created new Excel workbook (file not found): {output_file}")
        
        ws = wb.active
        ws.title = "Report"

        # Skip deletion routines if it's a new or reset workbook
        if not is_new_workbook:
            if delete_elevation_type:
                _delete_summary_section(ws)
                _delete_elevation_block(ws, delete_elevation_type, COL_A, PRICE_COL)
                _recalculate_running_grand_total(ws, PRICE_COL)
                _clean_trailing_blank_rows(ws, 1)
                try:
                    wb.save(output_file); create_summary_sheet(excel_path=output_file)
                except Exception as save_err:
                    print(f"❌ Error saving workbook during delete operation: {save_err}")
                    if completion_callback: completion_callback(f"Error saving report during delete: {save_err}")
                if completion_callback: completion_callback()
                return

            _delete_summary_section(ws)
            if not reset and elevation_type: _delete_elevation_block(ws, elevation_type, COL_A, PRICE_COL)
            
            _clean_trailing_blank_rows(ws, 1)
        
        start_row_for_new_block = 1
        if reset: # If reset, explicitly clear the worksheet. This is already handled by new Workbook creation for 'mode=new'
            # If wb was loaded and reset is true, we need to clear it. If it was a new workbook, it's already empty.
            if not is_new_workbook: # Only clear if we loaded an existing workbook and then reset it
                ws.delete_rows(1, ws.max_row) 
            print("🧹 Worksheet cleared for new report.")
        elif ws.max_row > 0 and not is_new_workbook: # Only find last row if it's an existing workbook with content
            last_gt_row = _find_row_by_value(ws, PRICE_COL, "RUNNING GRAND TOTAL", reverse=True)
            start_row_for_new_block = (last_gt_row + 3) if last_gt_row else (ws.max_row + 2)
        else: # For a truly new workbook or after a reset of an existing one, start from row 1
            start_row_for_new_block = 1


        input_data = [
            ("System Input", system_input), ("Elevation Type", elevation_type), ("Total Count", total_count),
            ("# Bays Wide", bays_wide), ("# Bays Tall", bays_tall), ("Opening Width", opening_width),
            ("Opening Height", opening_height), ("Sq Ft per Type", sqft_per_type), ("Total Sq Ft", total_sqft),
            ("Perimeter Ft", perimeter_ft), ("Total Perimeter Ft", total_perimeter_ft)
        ]
        for i, (header, value) in enumerate(input_data):
            ws.cell(row=start_row_for_new_block + i, column=COL_A, value=header).font = Font(bold=True)
            ws.cell(row=start_row_for_new_block + i, column=COL_B, value=value)

        current_system_total = [0.0]
        # This block should execute whether it's a new or existing report
        if all_elevations: # Renamed from `all_elevations and not reset` to ensure output for new reports
            profiles, accessories, other_manual_outputs = [], [], []
            for item in calculated_outputs:
                pn = item.get('part_number')
                manual = item.get('manual', False) # Get the manual flag

                if manual: # If the item is explicitly manual, group it with other manual outputs
                    other_manual_outputs.append(item)
                elif pn and pn != "N/A": # Otherwise, categorize based on part number mapping
                    if pn in PART_NUMBER_MAP.get("profiles", []): profiles.append(item)
                    elif pn in PART_NUMBER_MAP.get("accessories", []): accessories.append(item)
                    else: other_manual_outputs.append(item) # Fallback for non-manual items not in profiles/accessories
                else: # Items without a part number and not explicitly manual (should ideally be handled by 'manual' flag)
                    other_manual_outputs.append(item)

            output_section_current_row = start_row_for_new_block
            output_section_current_row = _write_output_section(ws, "PROFILES", profiles, COL_E, multiplier, current_system_total, output_section_current_row)
            output_section_current_row = _write_output_section(ws, "ACCESSORIES", accessories, COL_E, multiplier, current_system_total, output_section_current_row)

            grouped_other = {}; 
            for item in other_manual_outputs:
                # Use the 'type' field from the item for grouping, defaulting to 'MANUAL ITEMS'
                grouped_other.setdefault(item.get('type', 'MANUAL ITEMS').upper(), []).append(item)
            
            for grp_title, grp_items in grouped_other.items():
                output_section_current_row = _write_output_section(ws, grp_title, grp_items, COL_E, 1.0, current_system_total, output_section_current_row)

        system_total_row = ws.max_row + 2
        ws.cell(row=system_total_row, column=PRICE_COL, value="SYSTEM TOTAL").font = Font(bold=True)
        ws.cell(row=system_total_row + 1, column=PRICE_COL, value=current_system_total[0]).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        print(f"📊 System Total for '{elevation_type}': ${current_system_total[0]:.2f}")

        _recalculate_running_grand_total(ws, PRICE_COL)
        _autofit_columns(ws, COL_A, PRICE_COL, 1, ws.max_row)
        _clean_trailing_blank_rows(ws, 1)
        try:
            wb.save(output_file); print(f"✅ Excel report '{output_file}' updated successfully.")
        except Exception as save_err:
            print(f"❌ Error saving Excel report to '{output_file}': {save_err}")
            if completion_callback: completion_callback(f"Error saving report: {save_err}")
            return

        create_summary_sheet(excel_path=output_file)
        if completion_callback: completion_callback()
    except Exception as e:
        print(f"❌ An unexpected error occurred during Excel report generation: {e}")
        if completion_callback:
            completion_callback(f"Error generating report: {e}")