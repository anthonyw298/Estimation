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
        qty, pn = item.get('quantity', 0), item.get('part_number')
        unit_price, unit_type = (item.get('price', 0.0), item.get('unit', 'pcs')) if not pn or pn == "N/A" else get_price_by_part(pn)
        unit_price = (unit_price or 0.0) * (multiplier if title == "PROFILES" else 1.0)
        total_price = qty * unit_price; system_total_ref[0] += total_price

        ws.cell(row=current_row, column=colE, value=item.get('description', ''))
        ws.cell(row=current_row, column=colE + 1, value=pn or 'N/A')
        ws.cell(row=current_row, column=colE + 2, value=f"{qty} {unit_type}")
        ws.cell(row=current_row, column=colE + 3, value=total_price).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        current_row += 1
    return current_row + 1

def _delete_summary_section(ws):
    """Deletes the existing summary section from the worksheet."""
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
    """Reads elevation data, aggregates quantities/prices, and writes a clean summary."""
    try: data = json.load(open(json_path, 'r'))
    except (FileNotFoundError, json.JSONDecodeError) as e: print(f"⚠️ Error loading JSON file '{json_path}': {e}. Skipping summary."); return

    try: wb = load_workbook(excel_path)
    except FileNotFoundError: print(f"⚠️ Excel file '{excel_path}' not found. Cannot update summary."); return
    ws = wb.active
    
    _delete_summary_section(ws)
    if len(data) <= 1: wb.save(excel_path); print("ℹ️ Only one or zero elevations found. No summary sheet created/updated."); return

    summary = {}
    for elev_data in data.values():
        for output in elev_data.get('calculated_outputs', []):
            pn, desc, qty, price_manual = output.get('part_number'), output.get('description', '').strip(), output.get('quantity', 0), output.get('price')
            is_manual = not pn or pn == "N/A"
            key = desc if is_manual else pn
            if key not in summary: summary[key] = {'quantity': 0, 'price_manual': price_manual, 'is_manual': is_manual, 'description': desc}
            summary[key]['quantity'] += qty

    summary_rows = []
    for key, entry in summary.items():
        qty = entry['quantity']
        price_per_unit = (entry['price_manual'] or 0.0) if entry['is_manual'] else (get_price_by_part(key)[0] or 0.0)
        summary_rows.append((key, qty, price_per_unit * qty))

    last_gt_row = _find_row_by_value(ws, 8, "RUNNING GRAND TOTAL", reverse=True)
    start_row = (last_gt_row + 3) if last_gt_row else (ws.max_row + 2)

    ws.cell(row=start_row, column=1, value="Part Number / Description").font = Font(bold=True)
    ws.cell(row=start_row, column=2, value="Total Quantity").font = Font(bold=True)
    ws.cell(row=start_row, column=3, value="Total Price").font = Font(bold=True)

    for idx, (key, qty, total) in enumerate(summary_rows, start=start_row + 1):
        ws.cell(row=idx, column=1, value=key)
        ws.cell(row=idx, column=2, value=qty)
        ws.cell(row=idx, column=3, value=total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    _autofit_columns(ws, 1, 3, start_row, start_row + len(summary_rows))
    _clean_trailing_blank_rows(ws, 1)
    wb.save(excel_path); print(f"✅ Summary sheet updated in {excel_path}.")

def generate_excel_report(
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, mode=None, reset=False, all_elevations=None,
    delete_elevation_type=None
):
    """Generates or updates an Excel report with elevation details and calculated outputs."""
    COL_A, COL_B, COL_E, PRICE_COL = 1, 2, 5, 8
    multiplier = {"clear": 1.0, "black": 1.1, "paint": 1.2}.get(finish_input.lower(), 1.0)

    try: wb = load_workbook(output_file) if os.path.exists(output_file) and not (reset or mode == "new") else Workbook()
    except FileNotFoundError: wb = Workbook()
    ws = wb.active; ws.title = "Report"
    print(f"📄 {'Loaded' if os.path.exists(output_file) and not (reset or mode == 'new') else 'Created new'} Excel workbook: {output_file}")

    if delete_elevation_type:
        _delete_summary_section(ws)
        _delete_elevation_block(ws, delete_elevation_type, COL_A, PRICE_COL)
        _recalculate_running_grand_total(ws, PRICE_COL)
        _clean_trailing_blank_rows(ws, 1)
        wb.save(output_file); create_summary_sheet(excel_path=output_file)
        if completion_callback: completion_callback()
        return

    _delete_summary_section(ws)
    if not reset and elevation_type: _delete_elevation_block(ws, elevation_type, COL_A, PRICE_COL)
    
    _clean_trailing_blank_rows(ws, 1)
    start_row_for_new_block = 1
    if reset: ws.delete_rows(1, ws.max_row); print("🧹 Worksheet cleared for new report.")
    elif ws.max_row > 0:
        last_gt_row = _find_row_by_value(ws, PRICE_COL, "RUNNING GRAND TOTAL", reverse=True)
        start_row_for_new_block = (last_gt_row + 3) if last_gt_row else (ws.max_row + 2)

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
    if all_elevations and not reset:
        profiles, accessories, other_manual_outputs = [], [], []
        for item in calculated_outputs:
            pn, item_type = item.get('part_number'), item.get('type', '').lower()
            if pn and pn != "N/A":
                if pn in PART_NUMBER_MAP.get("profiles", []): profiles.append(item)
                elif pn in PART_NUMBER_MAP.get("accessories", []): accessories.append(item)
                else: other_manual_outputs.append(item)
            else: other_manual_outputs.append(item)

        output_section_current_row = start_row_for_new_block
        output_section_current_row = _write_output_section(ws, "PROFILES", profiles, COL_E, multiplier, current_system_total, output_section_current_row)
        output_section_current_row = _write_output_section(ws, "ACCESSORIES", accessories, COL_E, multiplier, current_system_total, output_section_current_row)

        grouped_other = {}; [grouped_other.setdefault(item.get('type', 'MANUAL ITEMS').upper(), []).append(item) for item in other_manual_outputs]
        for grp_title, grp_items in grouped_other.items():
            output_section_current_row = _write_output_section(ws, grp_title, grp_items, COL_E, 1.0, current_system_total, output_section_current_row)

    system_total_row = ws.max_row + 2
    ws.cell(row=system_total_row, column=PRICE_COL, value="SYSTEM TOTAL").font = Font(bold=True)
    ws.cell(row=system_total_row + 1, column=PRICE_COL, value=current_system_total[0]).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    print(f"📊 System Total for '{elevation_type}': ${current_system_total[0]:.2f}")

    _recalculate_running_grand_total(ws, PRICE_COL)
    _autofit_columns(ws, COL_A, PRICE_COL, 1, ws.max_row)
    _clean_trailing_blank_rows(ws, 1)
    wb.save(output_file); print(f"✅ Excel report '{output_file}' updated successfully.")
    create_summary_sheet(excel_path=output_file)
    if completion_callback: completion_callback()
