import os
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, numbers
from openpyxl.utils import get_column_letter
# Import the updated pricing functions
from utils.pricing import get_price_by_part, reverse_material_impact, load_extra_materials, save_extra_materials, apply_material_impact_to_extra_materials
from data.part_number import PART_NUMBER_MAP

output_file = "output.xlsx"
SAVED_ELEVATIONS_FILE = 'saved_elevations.json' # Define the JSON file path

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
            ws.delete_rows(current_row, 1)
            rows_deleted += 1
        else: current_row += 1
    if rows_deleted > 0: print(f"🧹 Cleaned {rows_deleted} trailing blank rows starting from row {start_row}.")

def _recalculate_running_grand_total(ws, price_col):
    """Recalculates and updates the 'RUNNING GRAND TOTAL' in the worksheet."""
    # Find and delete existing RUNNING GRAND TOTAL
    for r in range(ws.max_row, 0, -1):
        cell_value = ws.cell(row=r, column=price_col).value
        if isinstance(cell_value, str) and cell_value.strip() == "RUNNING GRAND TOTAL":
            # Delete RUNNING GRAND TOTAL and its value
            ws.delete_rows(r, 2)
            print("🗑️ Existing RUNNING GRAND TOTAL removed for recalculation.")
            break

    running_grand_total = 0.0
    last_system_total_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=price_col).value == "SYSTEM TOTAL":
            last_system_total_row = r
            val = ws.cell(row=r + 1, column=price_col).value
            if isinstance(val, (float, int)): running_grand_total += val
            elif isinstance(val, str) and val.strip().startswith("$"):
                try: running_grand_total += float(val.strip("$"))
                except ValueError: print(f"⚠️ Could not parse SYSTEM TOTAL value: {val}")
            else:
                print(f"⚠️ SYSTEM TOTAL value at row {r+1}, column {price_col} is not a number: {val}")

    new_gt_row = (last_system_total_row + 3) if last_system_total_row else (ws.max_row + 2)
    
    # Only write RUNNING GRAND TOTAL if there are any system totals
    if running_grand_total > 0 or last_system_total_row is not None:
        ws.cell(row=new_gt_row, column=price_col, value="RUNNING GRAND TOTAL").font = Font(bold=True)
        ws.cell(row=new_gt_row + 1, column=price_col, value=running_grand_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        print(f"📈 Running Grand Total recalculated and updated to: ${running_grand_total:.2f}.")
    else:
        print("ℹ️ No system totals found. RUNNING GRAND TOTAL not written.")


def _write_output_section(ws, title, items, colE, multiplier, system_total_ref, start_output_row, current_extra_materials_state):
    """
    Writes a section of calculated outputs (e.g., PROFILES, ACCESSORIES) to the worksheet.
    Returns the next available row and the material impacts for the current section.
    `current_extra_materials_state` is passed to get_price_by_part for simulation.
    """
    if not items: return start_output_row, [] # Return current row + 1 and empty impacts

    current_row = start_output_row # Start section at the provided start_output_row
    
    ws.cell(row=current_row, column=colE, value=title).font = Font(bold=True)
    for i, h in enumerate(["Description", "Part Number", "Quantity", "Price"]):
        ws.cell(row=current_row + 1, column=colE + i, value=h).font = Font(bold=True)
    current_row += 2

    section_material_impacts = [] # This will collect impacts for this section

    for item in items:
        qty = item.get('quantity', 0)
        pn = item.get('part_number')
        manual = item.get('manual', False)

        total_item_price = 0.0
        unit_type = "pcs"
        material_impact_details = None # Initialize to None

        if manual: # If it's explicitly marked as manual
            if pn and pn != "N/A": # If manual and has a part number, try to get price from part data
                # Pass current_extra_materials_state for accurate simulation AND group=True for length-based pricing
                price_calculated, unit_calculated, material_impact_details = \
                    get_price_by_part(pn, qty, current_extra_materials=current_extra_materials_state, summary=False, group=True) 
                # For manual, if get_price_by_part returns a valid price, use it, otherwise fall back to 'price' in item
                total_item_price = (price_calculated if price_calculated is not None else item.get('price', 0.0) * qty)
                unit_type = unit_calculated or item.get('unit', 'pcs')
            else: # Manual entry without a part number (or "N/A")
                total_item_price = item.get('price', 0.0) * qty
                unit_type = item.get('unit', 'pcs')
                # For pure manual items without PN, create a generic impact for tracking if needed for summary
                material_impact_details = {
                    'part_number': "N/A - Manual", # Use a distinct PN for manual items without real PN
                    'requested_qty': qty,
                    'purchased_qty_or_length': 0.0,
                    'leftover_generated_qty_or_length': 0.0,
                    'used_from_leftover_qty_or_length': 0.0,
                    'cost_incurred': total_item_price
                }
        else: # Not a manual entry, always get price from part data
            # get_price_by_part now returns the total cost for the requested qty and unit_type, plus material_impact_details
            total_item_price, unit_type, material_impact_details = \
                get_price_by_part(pn, qty, current_extra_materials=current_extra_materials_state, summary=False) # Not summary here
            total_item_price = total_item_price or 0.0
            unit_type = unit_type or "pcs"
        
        # Add a check for None in material_impact_details for robustness
        if material_impact_details:
             section_material_impacts.append(material_impact_details)

        # Apply multiplier for profiles to the total calculated item price
        if title == "PROFILES":
            total_item_price *= multiplier

        system_total_ref[0] += total_item_price

        ws.cell(row=current_row, column=colE, value=item.get('description', ''))
        ws.cell(row=current_row, column=colE + 1, value=pn or 'N/A')
        ws.cell(row=current_row, column=colE + 2, value=f"{qty} {unit_type}")
        ws.cell(row=current_row, column=colE + 3, value=total_item_price).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        current_row += 1
    return current_row + 1, section_material_impacts


def _delete_summary_section(ws):
    """Deletes the existing summary section from the worksheet."""
    # Only attempt to delete if there's more than one row (i.e., not a completely empty sheet)
    if ws.max_row <= 1: # A new sheet often has 1 row by default, which is empty.
        print("ℹ️ Worksheet is largely empty. Skipping summary section deletion.")
        return

    summary_start_row = _find_row_by_value(ws, 1, "Part Number / Description")
    if summary_start_row:
        current_row_to_delete = summary_start_row
        # Find the end of the summary section (first blank row after header or end of sheet)
        while current_row_to_delete <= ws.max_row:
            # Check if the row is entirely empty or if it's the start of another section
            # A simple check for the first column being None is usually sufficient
            if ws.cell(row=current_row_to_delete, column=1).value is None and \
               ws.cell(row=current_row_to_delete, column=2).value is None and \
               ws.cell(row=current_row_to_delete, column=3).value is None:
                break
            current_row_to_delete += 1
        
        if current_row_to_delete > summary_start_row:
            rows_to_delete = current_row_to_delete - summary_start_row
            ws.delete_rows(summary_start_row, rows_to_delete)
            print(f"🗑️ Existing summary section cleared ({rows_to_delete} rows) starting at row {summary_start_row}.");
            _clean_trailing_blank_rows(ws, summary_start_row)
        else: print("ℹ️ Summary header found but no data rows to delete.")
    else: print("ℹ️ No existing summary section found to delete.")

# --- Main Functions ---

def create_summary_sheet(excel_path=output_file, json_path=SAVED_ELEVATIONS_FILE):
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
        ws.title = "Report"
    except Exception as e:
        print(f"⚠️ Excel file '{excel_path}' not found or corrupted for summary: {e}. Cannot update summary sheet.")
        return
    
    _delete_summary_section(ws) 
    
    if len(data) == 0: # If no elevations, just clear summary and save
        try:
            if wb: wb.save(excel_path)
            print("ℹ️ No elevations found. Summary sheet cleared (if existed) and not re-created.")
        except Exception as save_err:
            print(f"❌ Error saving workbook after no summary data: {save_err}")
        return

# === FIRST: Aggregate all quantities by part number or description ===
    aggregated_summary = {}
    # No need to load extra_materials here, as get_price_by_part with summary=True will ignore it.

    for elev_data in data.values():
        for output in elev_data.get('calculated_outputs', []):
            part_number = output.get('part_number', '').strip()
            description = output.get('description', '').strip()
            quantity = output.get('quantity', 0)
            manual = output.get('manual', False)

            if manual or not part_number or part_number == "N/A":
                key = f"MANUAL_{description}_{part_number}" # Unique key for manual items
                display_key_base = description
                if manual and part_number and part_number != "N/A":
                    display_key_base = f"{description} ({part_number})"

            else:
                key = part_number
                display_key_base = part_number

            if key not in aggregated_summary:
                aggregated_summary[key] = {
                    'quantity': 0,
                    'description': description, # Original description
                    'display_key': display_key_base, # Key for display in summary
                    'part_number': part_number,
                    'manual': manual,
                    'price': output.get('price', 0.0), # Stored per-unit price for manual calculations
                    'unit': output.get('unit', 'pcs') # Original unit from output
                }

            aggregated_summary[key]['quantity'] += quantity

    # === SECOND: For each unique item, run get_price_by_part for final cost ===
    final_summary_data = []

    for key, item in aggregated_summary.items():
        quantity = item['quantity']
        manual = item['manual']
        part_number = item['part_number']
        display_key = item['display_key'] # Use the prepared display key
        original_unit = item['unit']

        total_cost_for_item = 0.0
        calculated_unit_type = original_unit # Default unit for display

        if manual:
            if part_number and part_number != "N/A":
                # For manual items with PN, always use get_price_by_part for the cost.
                # If get_price_by_part returns None, it means the part number wasn't found in the database.
                price_from_part, unit_type_from_pricing, _ = \
                    get_price_by_part(part_number, quantity, summary=True, group=True) # Pass group=True here
                
                total_cost_for_item = price_from_part if price_from_part is not None else 0.0 # If part not found, cost is 0
                calculated_unit_type = unit_type_from_pricing or original_unit
            else:
                # For truly manual entries without a part number, use the stored price * quantity
                total_cost_for_item = item['price'] * quantity
                calculated_unit_type = original_unit # Use original unit for pure manual

            final_summary_data.append((display_key, f"{quantity} {calculated_unit_type}", total_cost_for_item))
        else: # Auto items
            total_price, unit_type_from_pricing, _ = \
                get_price_by_part(part_number, quantity, summary=True) # Pass summary=True
            total_cost_for_item = total_price if total_price is not None else 0.0
            calculated_unit_type = unit_type_from_pricing or original_unit # Get from pricing, fallback to original

            final_summary_data.append((display_key, f"{quantity} {calculated_unit_type}", total_cost_for_item))

    last_gt_row = _find_row_by_value(ws, 8, "RUNNING GRAND TOTAL", reverse=True)
    start_row = (last_gt_row + 3) if last_gt_row else (ws.max_row + 2)

    if not final_summary_data: # If no data to summarize, don't write headers
        print("ℹ️ No data to summarize. Summary section not written.")
        try:
            wb.save(excel_path)
        except Exception as save_err:
            print(f"❌ Error saving workbook after no summary data: {save_err}")
        return

    ws.cell(row=start_row, column=1, value="Part Number / Description").font = Font(bold=True)
    ws.cell(row=start_row, column=2, value="Total Quantity").font = Font(bold=True)
    ws.cell(row=start_row, column=3, value="Total Price").font = Font(bold=True)

    for idx, (item_key, qty_with_unit, total_cost) in enumerate(final_summary_data, start=start_row + 1):
        ws.cell(row=idx, column=1, value=item_key)
        ws.cell(row=idx, column=2, value=qty_with_unit)
        price_cell = ws.cell(row=idx, column=3, value=total_cost)
        price_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    _autofit_columns(ws, 1, 3, start_row, start_row + len(final_summary_data))
    _clean_trailing_blank_rows(ws, 1)
    try:
        wb.save(excel_path); print(f"✅ Summary sheet updated in {excel_path}.")
    except Exception as save_err:
        print(f"❌ Error saving summary sheet to '{excel_path}': {save_err}")

def generate_excel_report(
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, reset=False, delete_elevation_type=None,
    door_size=None # Added door_size as an optional parameter
):
    """Generates or updates an Excel report with detailed elevation inputs and calculated outputs."""
    COL_A, COL_B, COL_E, PRICE_COL = 1, 2, 5, 8
    
    # --- Always load current saved elevations from JSON at the start ---
    current_saved_elevations = {}
    if os.path.exists(SAVED_ELEVATIONS_FILE):
        try:
            with open(SAVED_ELEVATIONS_FILE, 'r') as f:
                current_saved_elevations = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"❌ Error loading {SAVED_ELEVATIONS_FILE}: {e}. Starting with empty elevations in memory.")

    # --- Handle Deletion First (if requested) ---
    if delete_elevation_type:
        print(f"🚀 Processing deletion request for '{delete_elevation_type}'.")
        
        # 1. Get the material_impact data for the elevation being deleted from memory
        elevation_to_delete_data = current_saved_elevations.get(delete_elevation_type)
        if elevation_to_delete_data and 'material_impact' in elevation_to_delete_data:
            print(f"Attempting to reverse material impact for '{delete_elevation_type}'.")
            reverse_material_impact(elevation_to_delete_data['material_impact'])
            print(f"✅ Material impact for '{delete_elevation_type}' reversed in extra_materials.json.")
        else:
            print(f"ℹ️ No material impact data found for '{delete_elevation_type}' in saved data to reverse.")

        # 2. Delete the elevation from saved_elevations.json (in memory)
        if delete_elevation_type in current_saved_elevations:
            del current_saved_elevations[delete_elevation_type]
            print(f"Deleted '{delete_elevation_type}' from saved_elevations.json in memory.")
        else:
            print(f"⚠️ '{delete_elevation_type}' not found in {SAVED_ELEVATIONS_FILE}. Nothing to delete from JSON.")

        # 3. Save the updated saved_elevations.json to persist the deletion
        try:
            with open(SAVED_ELEVATIONS_FILE, 'w') as f:
                json.dump(current_saved_elevations, f, indent=4)
            print(f"✅ {SAVED_ELEVATIONS_FILE} updated after deletion.")
        except IOError as e:
            print(f"❌ Error saving updated {SAVED_ELEVATIONS_FILE} during delete: {e}")
            if completion_callback: completion_callback(f"Error saving updated elevations after delete: {e}")
            return # Critical error, cannot proceed

        # After deletion, proceed to rebuild the entire Excel report from the remaining elevations
        print("🔄 Rebuilding entire Excel report after elevation deletion...")
        
    else: # Handle New/Update Elevation
        # If this elevation already exists in saved_elevations.json, reverse its old impact first.
        if elevation_type in current_saved_elevations and not reset: 
            print(f"🔄 Detected update for '{elevation_type}'. Reversing old material impact before applying new.")
            old_elevation_data = current_saved_elevations[elevation_type]
            if 'material_impact' in old_elevation_data:
                reverse_material_impact(old_elevation_data['material_impact'])
            else:
                print(f"ℹ️ No old material impact data found for '{elevation_type}' to reverse (might be a very old entry).")
        
        # Update the current elevation data in memory
        current_saved_elevations[elevation_type] = {
            "system": system_input,
            "finish": finish_input,
            "total_count": total_count,
            "bays_wide": bays_wide,
            "bays_tall": bays_tall,
            "opening_width_inches": opening_width,
            "opening_height_inches": opening_height,
            "sqft_per_type": sqft_per_type,
            "total_sqft": total_sqft,
            "perimeter_ft": perimeter_ft,
            "total_perimeter_ft": total_perimeter_ft,
            "calculated_outputs": calculated_outputs,
            "material_impact": [] # This will be populated during the rebuild process
        }
        if door_size is not None:
            current_saved_elevations[elevation_type]['door_size'] = door_size

        # Save the updated saved_elevations.json to persist the add/update
        try:
            with open(SAVED_ELEVATIONS_FILE, 'w') as f:
                json.dump(current_saved_elevations, f, indent=4)
            print(f"✅ Elevation '{elevation_type}' saved to {SAVED_ELEVATIONS_FILE}.")
        except IOError as e:
            print(f"❌ Error saving elevation to {SAVED_ELEVATIONS_FILE}: {e}")
            if completion_callback: completion_callback(f"Error saving elevation: {e}")
            return

    # --- Consolidated Excel Rebuild Logic ---
    wb = Workbook() # Always start with a completely new workbook for rebuild
    ws = wb.active
    ws.title = "Report"
    
    # Reset extra_materials.json for a clean rebuild process.
    # This is CRUCIAL to ensure that all remaining elevations are priced
    # and impact extra_materials.json as if they were added sequentially
    # starting from an empty material state.
    save_extra_materials({}) # Wipe extra_materials.json
    overall_current_extra_materials_state = load_extra_materials() # Reload as empty for first calculation

    current_excel_row = 1
    
    # To maintain a stable order in Excel, get sorted keys
    sorted_elev_names = sorted(current_saved_elevations.keys())

    if not sorted_elev_names:
        print("ℹ️ No elevations remaining. Excel report will be empty.")
        _clean_trailing_blank_rows(ws, 1) # Clean any initial blank rows
    else:
        for elev_name in sorted_elev_names:
            elev_data = current_saved_elevations[elev_name]
            print(f"🛠️ Re-adding '{elev_name}' to the report during rebuild...")
            
            # Extract all necessary parameters from `elev_data`
            current_system_input = elev_data.get("system")
            current_finish_input = elev_data.get("finish")
            current_total_count = elev_data.get("total_count")
            current_bays_wide = elev_data.get("bays_wide")
            current_bays_tall = elev_data.get("bays_tall")
            current_opening_width = elev_data.get("opening_width_inches")
            current_opening_height = elev_data.get("opening_height_inches")
            current_sqft_per_type = elev_data.get("sqft_per_type")
            current_total_sqft = elev_data.get("total_sqft")
            current_perimeter_ft = elev_data.get("perimeter_ft")
            current_total_perimeter_ft = elev_data.get("total_perimeter_ft")
            current_calculated_outputs = elev_data.get("calculated_outputs")
            current_door_size = elev_data.get("door_size") # Get door_size from saved data

            multiplier = {"clear": 1.0, "black": 1.1, "paint": 1.2}.get(current_finish_input.lower(), 1.0)
            
            system_total_for_this_block = [0.0]
            newly_calculated_material_impacts_for_this_elevation = []

            # Write elevation input data for this block
            input_data = [
                ("System Input", current_system_input), ("Elevation Type", elev_name), ("Total Count", current_total_count),
                ("Bays Wide", current_bays_wide), ("Bays Tall", current_bays_tall), ("Opening Width", current_opening_width),
                ("Opening Height", current_opening_height), ("Sq Ft per Type", current_sqft_per_type), ("Total Sq Ft", current_total_sqft),
                ("Perimeter Ft", current_perimeter_ft), ("Total Perimeter Ft", current_total_perimeter_ft)
            ]
            if current_door_size is not None: # Add door_size to input_data if it exists
                input_data.append(("Door Size", current_door_size))

            for i, (header, value) in enumerate(input_data):
                ws.cell(row=current_excel_row + i, column=COL_A, value=header).font = Font(bold=True)
                ws.cell(row=current_excel_row + i, column=COL_B, value=value)
                # Column C remains empty here, providing visual separation before COL_E

            # Calculate the starting row for the output sections
            # This ensures "PROFILES" starts directly on the same row as the input data, in column E
            output_section_current_row = current_excel_row 
            
            # Initialize lists for outputs before populating them to prevent UnboundLocalError
            profiles_for_section = []
            accessories_for_section = []
            other_manual_outputs_for_section = []

            for item in current_calculated_outputs:
                pn = item.get('part_number')
                manual = item.get('manual', False)

                if manual:
                    other_manual_outputs_for_section.append(item)
                elif pn and pn != "N/A":
                    if pn in PART_NUMBER_MAP.get("profiles", []): profiles_for_section.append(item)
                    elif pn in PART_NUMBER_MAP.get("accessories", []): accessories_for_section.append(item)
                    else: other_manual_outputs_for_section.append(item)
                else:
                    other_manual_outputs_for_section.append(item)
            
            # Write output sections and collect material impacts
            # Pass the output_section_current_row and update it for subsequent sections
            next_row_after_profiles, impacts = _write_output_section(ws, "PROFILES", profiles_for_section, COL_E, multiplier, system_total_for_this_block, output_section_current_row, overall_current_extra_materials_state)
            newly_calculated_material_impacts_for_this_elevation.extend(impacts)
            
            next_row_after_accessories, impacts = _write_output_section(ws, "ACCESSORIES", accessories_for_section, COL_E, multiplier, system_total_for_this_block, next_row_after_profiles, overall_current_extra_materials_state)
            newly_calculated_material_impacts_for_this_elevation.extend(impacts)
            
            grouped_other = {}; 
            for item in other_manual_outputs_for_section:
                grouped_other.setdefault(item.get('type', 'MANUAL ITEMS').upper(), []).append(item)
            
            current_section_row = next_row_after_accessories
            for grp_title, grp_items in grouped_other.items():
                next_row_after_group, impacts = _write_output_section(ws, grp_title, grp_items, COL_E, 1.0, system_total_for_this_block, current_section_row, overall_current_extra_materials_state)
                newly_calculated_material_impacts_for_this_elevation.extend(impacts)
                current_section_row = next_row_after_group # Update for the next group

            # After collecting all impacts for this elevation, apply them to extra_materials.json
            # This happens *after* all calculations for the elevation are done and impacts are collected.
            for impact_detail in newly_calculated_material_impacts_for_this_elevation:
                apply_material_impact_to_extra_materials(impact_detail)
            print(f"✅ Material impacts for '{elev_name}' applied during rebuild.")

            # IMPORTANT: Reload the overall_current_extra_materials_state after applying impacts for the current elevation.
            # This ensures the next iteration of the loop (for the next elevation) uses the most up-to-date inventory.
            overall_current_extra_materials_state = load_extra_materials()
            print(f"🔄 Reloaded extra_materials for next elevation: {overall_current_extra_materials_state}")


            # Update the `material_impact` in the in-memory `current_saved_elevations` for this elevation
            # with the newly calculated impacts during this rebuild phase. This is important
            # if you ever delete THIS elevation later.
            current_saved_elevations[elev_name]['material_impact'] = newly_calculated_material_impacts_for_this_elevation

            # Write SYSTEM TOTAL for this elevation
            system_total_row = ws.max_row + 2 
            ws.cell(row=system_total_row, column=PRICE_COL, value="SYSTEM TOTAL").font = Font(bold=True)
            ws.cell(row=system_total_row + 1, column=PRICE_COL, value=system_total_for_this_block[0]).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            print(f"📊 Rebuilt System Total for '{elev_name}': ${system_total_for_this_block[0]:.2f}")
            
            current_excel_row = system_total_row + 3 # Move to the next block start

    # After rebuilding all elevations
    _recalculate_running_grand_total(ws, PRICE_COL) # Recalculate based on the rebuilt sheet
    _autofit_columns(ws, COL_A, PRICE_COL, 1, ws.max_row)
    _clean_trailing_blank_rows(ws, 1)
    
    try:
        wb.save(output_file)
        print(f"✅ Excel report '{output_file}' fully rebuilt.")
    except Exception as save_err:
        print(f"❌ Error saving Excel report during full rebuild: {save_err}")
        if completion_callback: completion_callback(f"Error saving report: {save_err}")
        return

    # Save the `current_saved_elevations` again to ensure the `material_impact` field is updated
    # with the newly calculated impacts during the rebuild process.
    try:
        with open(SAVED_ELEVATIONS_FILE, 'w') as f:
            json.dump(current_saved_elevations, f, indent=4)
        print(f"✅ All elevations in {SAVED_ELEVATIONS_FILE} updated with rebuilt material impacts.")
    except IOError as e:
        print(f"❌ Error saving all elevations to {SAVED_ELEVATIONS_FILE} after rebuild: {e}")

    create_summary_sheet(excel_path=output_file) # Rebuild summary based on the final JSON
    if completion_callback: completion_callback()
