import os
import json
from openpyxl import Workbook
from openpyxl.styles import Font, numbers
from openpyxl.utils import get_column_letter

# Removed global constants for file paths, as they will now be passed as arguments
from utils.pricing import get_price_by_part, reverse_material_impact, load_extra_materials, save_extra_materials, apply_material_impact_to_extra_materials_in_memory, get_unit_price_by_part
from data.part_number import PART_NUMBER_MAP

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
    end_row = end_row if end_row is not None else ws.max_row
    for col_idx in range(start_col, end_col + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for r in range(start_row, end_row + 1):
            cell_value = ws.cell(row=r, column=col_idx).value
            if cell_value is not None:
                max_len = max(max_len, len(str(cell_value)))
        current_width_obj = ws.column_dimensions[col_letter]
        current_width = current_width_obj.width if current_width_obj.width is not None else 0.0
        if col_idx == 5: # Column E (Description)
            if max_len > current_width:
                ws.column_dimensions[col_letter].width = max_len
        else: # For all other columns
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
    if rows_deleted > 0: print(f"Cleaned {rows_deleted} trailing blank rows starting from row {start_row}.")

def _recalculate_running_grand_total(ws, price_col):
    """Recalculates and updates the 'RUNNING GRAND TOTAL' in the worksheet."""
    for r in range(ws.max_row, 0, -1):
        if isinstance(ws.cell(row=r, column=price_col).value, str) and ws.cell(row=r, column=price_col).value.strip() == "RUNNING GRAND TOTAL":
            ws.delete_rows(r, 2)
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
                except ValueError: pass

    new_gt_row = (last_system_total_row + 3) if last_system_total_row else (ws.max_row + 2)
    
    if running_grand_total > 0 or last_system_total_row is not None:
        ws.cell(row=new_gt_row, column=price_col, value="RUNNING GRAND TOTAL").font = Font(bold=True)
        ws.cell(row=new_gt_row + 1, column=price_col, value=running_grand_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        print(f"Running Grand Total recalculated and updated to: ${running_grand_total:.2f}.")


def _write_output_section(ws, title, items, colE, multiplier, system_total_ref, start_output_row, current_extra_materials_state, extra_materials_path):
    """Writes a section of calculated outputs to the worksheet."""
    if not items: return start_output_row, []

    current_row = start_output_row
    ws.cell(row=current_row, column=colE, value=title).font = Font(bold=True)
    for i, h in enumerate(["Description", "Part Number", "Quantity", "Price"]):
        ws.cell(row=current_row + 1, column=colE + i, value=h).font = Font(bold=True)
    current_row += 2

    section_material_impacts = []

    for item in items:
        qty, pn, manual = item.get('quantity', 0), item.get('part_number'), item.get('manual', False)
        total_item_price, unit_type, material_impact_details = 0.0, "pcs", None

        if manual:
            if pn and pn != "N/A":
                price_calculated, unit_calculated, material_impact_details = \
                    get_price_by_part(pn, qty, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False, group=True) 
                total_item_price = (price_calculated if price_calculated is not None else item.get('price', 0.0) * qty)
                unit_type = unit_calculated or item.get('unit', 'pcs')
            else:
                total_item_price = item.get('price', 0.0) * qty
                unit_type = item.get('unit', 'pcs')
                material_impact_details = {
                    'part_number': "N/A - Manual", 'requested_qty': qty, 'purchased_qty_or_length': 0.0,
                    'leftover_generated_qty_or_length': 0.0, 'used_from_leftover_qty_or_length': 0.0,
                    'cost_incurred': total_item_price, 'type_processed_as': 'manual_no_pn'
                }
        else:
            total_price, unit_type, material_impact_details = \
                get_price_by_part(pn, qty, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False)
            total_item_price = total_price or 0.0
            unit_type = unit_type or "pcs"
        
        if material_impact_details:
            section_material_impacts.append(material_impact_details)
            apply_material_impact_to_extra_materials_in_memory(current_extra_materials_state, material_impact_details)

        if title == "PROFILES": total_item_price *= multiplier
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

def create_summary_sheet(excel_path, elevations_json_path, extra_materials_json_path):
    """Reads elevation data, aggregates quantities and prices by part number,
    and writes a clean summary section into the Excel file, including reusable material data."""

    import json
    from openpyxl import load_workbook
    from openpyxl.styles import Font, numbers

    # === Load elevations data ===
    try:
        data = json.load(open(elevations_json_path, 'r'))
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"⚠️ Error loading JSON file '{elevations_json_path}': {e}. Skipping summary.")
        return

    # === Load extra materials properly ===
    try:
        extra_materials = load_extra_materials(extra_materials_json_path) # Pass the path here
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"⚠️ Error loading Extra Materials file '{extra_materials_json_path}': {e}. Reusable data will be skipped.")
        extra_materials = {}

    # === Load Excel workbook ===
    try:
        wb = load_workbook(excel_path)
        ws = wb.active
        ws.title = "Report"
    except Exception as e:
        print(f"⚠️ Excel file '{excel_path}' not found or corrupted for summary: {e}. Cannot update summary sheet.")
        return

    _delete_summary_section(ws)

    if not data:  # Nothing to summarize
        try:
            wb.save(excel_path)
            print("ℹ️ No elevations found. Summary sheet cleared (if existed) and not re-created.")
        except Exception as save_err:
            print(f"❌ Error saving workbook after no summary data: {save_err}")
        return

    # === Aggregate ===
    aggregated_summary = {}

    for elev_data in data.values():
        for output in elev_data.get('calculated_outputs', []):
            part_number = output.get('part_number', '').strip()
            description = output.get('description', '').strip()
            quantity = output.get('quantity', 0)
            manual = output.get('manual', False)

            # Determine the key for aggregation and display_key_base
            if manual:
                # For manual items, group by description and part number (if exists)
                key = f"MANUAL_{description}_{part_number}"
                display_key_base = description
                if part_number and part_number != "N/A":
                    display_key_base = f"{description} ({part_number})"
            else:
                # For non-manual items, group by part number
                key = part_number
                display_key_base = part_number

            if key not in aggregated_summary:
                aggregated_summary[key] = {
                    'quantity': 0,
                    'description': description,
                    'display_key': display_key_base,
                    'part_number': part_number,
                    'manual': manual,
                    'price': output.get('price', 0.0), # Store the individual item price, not total
                    'unit': output.get('unit', 'pcs')
                }

            # Safely add quantities: convert to float or 0 if invalid
            try:
                qty_float = float(quantity)
            except (TypeError, ValueError):
                qty_float = 0.0

            aggregated_summary[key]['quantity'] += qty_float

    # === Final summary rows ===
    final_summary_data = []

    for key, item in aggregated_summary.items():
        quantity = item['quantity']
        manual = item['manual']
        part_number = item['part_number']
        display_key = item['display_key']
        original_unit = item['unit']

        total_cost_for_item = 0.0
        calculated_unit_type = original_unit
        reusable_qty_f = 0.0
        reusable_percentage = 0.0
        reusable_cost_saved_val = 0.0
        unit_for_reusable_qty_display = original_unit # Default to original unit

        # Calculate total_cost_for_item and calculated_unit_type
        if manual:
            if part_number and part_number != "N/A":
                # For manual items with a part number, get price from pricing module
                price_from_part, unit_type_from_pricing, _ = get_price_by_part(part_number, quantity, extra_materials_file=extra_materials_json_path, summary=True, group=True)
                total_cost_for_item = price_from_part if price_from_part is not None else 0.0
                calculated_unit_type = unit_type_from_pricing or original_unit
            else:
                # For manual items without a part number, use the provided price * quantity
                try:
                    price = float(item['price'])
                except (TypeError, ValueError):
                    price = 0.0
                try:
                    qty_float = float(quantity)
                except (TypeError, ValueError):
                    qty_float = 0.0
                total_cost_for_item = price * qty_float
                calculated_unit_type = original_unit
        else:
            # For non-manual items, get price from pricing module
            total_price, unit_type_from_pricing, _ = get_price_by_part(part_number, quantity, extra_materials_file=extra_materials_json_path, summary=True)
            total_cost_for_item = total_price if total_price is not None else 0.0
            calculated_unit_type = unit_type_from_pricing or original_unit

        # Calculate reusable material data for items with a part number (manual or non-manual)
        if part_number and part_number != "N/A":
            part_data = extra_materials.get(part_number, {})
            reusable_qty = part_data.get("quantity", 0)
            length_pieces = part_data.get("length_pieces", [])

            if reusable_qty == 0 and length_pieces:
                try:
                    reusable_qty = sum(float(x) for x in length_pieces)
                except (TypeError, ValueError):
                    reusable_qty = 0.0

            try:
                reusable_qty_f = float(reusable_qty)
            except (TypeError, ValueError):
                reusable_qty_f = 0.0

            try:
                quantity_f = float(quantity)
            except (TypeError, ValueError):
                quantity_f = 0.0

            if quantity_f > 0:
                reusable_percentage = (reusable_qty_f / quantity_f) * 100

            unit_price, unit_type_for_reusable = get_unit_price_by_part(part_number, extra_materials_file=extra_materials_json_path)

            if reusable_qty_f > 0 and unit_price is not None:
                reusable_cost_saved_val = unit_price * reusable_qty_f
            else:
                reusable_cost_saved_val = 0.0
            
            unit_for_reusable_qty_display = unit_type_for_reusable or original_unit # Use unit from pricing if available

        final_summary_data.append((
            display_key,
            f"{quantity:.2f} {calculated_unit_type}", # Format quantity for display
            total_cost_for_item,
            reusable_qty_f,
            reusable_percentage,
            reusable_cost_saved_val,
            manual,
            unit_for_reusable_qty_display,
            part_number # Include part_number in the tuple for easier access
        ))


    last_gt_row = _find_row_by_value(ws, 8, "RUNNING GRAND TOTAL", reverse=True)
    start_row = (last_gt_row + 3) if last_gt_row else (ws.max_row + 2)

    if not final_summary_data:
        print("ℹ️ No data to summarize. Summary section not written.")
        try:
            wb.save(excel_path)
        except Exception as save_err:
            print(f"❌ Error saving workbook after no summary data: {save_err}")
        return

    ws.cell(row=start_row, column=1, value="Part Number / Description").font = Font(bold=True)
    ws.cell(row=start_row, column=2, value="Total Quantity").font = Font(bold=True)
    ws.cell(row=start_row, column=3, value="Total Price").font = Font(bold=True)
    ws.cell(row=start_row, column=4, value="Reusable Material Quantity").font = Font(bold=True)
    ws.cell(row=start_row, column=5, value="Reusable % of Total").font = Font(bold=True)
    ws.cell(row=start_row, column=6, value="Reusable Material Cost").font = Font(bold=True)

    for idx, (item_key, qty_with_unit, total_cost, reusable_qty, reusable_percentage, reusable_cost_saved_val, manual_flag, unit_for_reusable_qty_display, current_item_part_number) in enumerate(final_summary_data, start=start_row + 1):
        ws.cell(row=idx, column=1, value=item_key)
        ws.cell(row=idx, column=2, value=qty_with_unit)
        price_cell = ws.cell(row=idx, column=3, value=total_cost)
        price_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

        # Display reusable material info if a valid part number exists for the item
        if current_item_part_number and current_item_part_number != "N/A":
            ws.cell(row=idx, column=4, value=f"{reusable_qty:.2f} {unit_for_reusable_qty_display}")
            ws.cell(row=idx, column=5, value=f"{reusable_percentage:.2f}%").number_format = numbers.FORMAT_TEXT
            ws.cell(row=idx, column=6, value=reusable_cost_saved_val).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        else:
            ws.cell(row=idx, column=4, value="")
            ws.cell(row=idx, column=5, value="").number_format = numbers.FORMAT_TEXT
            ws.cell(row=idx, column=6, value="").number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    _autofit_columns(ws, 1, 6, start_row, start_row + len(final_summary_data))
    _clean_trailing_blank_rows(ws, 1)

    try:
        wb.save(excel_path)
        print(f"✅ Summary sheet updated in {excel_path}.")
    except Exception as save_err:
        print(f"❌ Error saving summary sheet to '{excel_path}': {save_err}")

def generate_excel_report(
    excel_path, elevations_json_path, extra_materials_json_path, # New parameters
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, reset=False, delete_elevation_type=None,
    door_size=None, mode=None # Added mode to handle "export_all" from main.py
):
    """Generates or updates an Excel report with detailed elevation inputs and calculated outputs."""
    COL_A, COL_B, COL_E, PRICE_COL = 1, 2, 5, 8
    
    current_saved_elevations = {}
    if os.path.exists(elevations_json_path): # Use elevations_json_path
        try:
            with open(elevations_json_path, 'r') as f:
                current_saved_elevations = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Error loading {elevations_json_path}: {e}. Starting with empty elevations in memory.")

    # --- Handle Deletion or Update/Save Mode ---
    if delete_elevation_type:
        elevation_to_delete_data = current_saved_elevations.get(delete_elevation_type)
        if elevation_to_delete_data and 'material_impact' in elevation_to_delete_data:
            reverse_material_impact(elevation_to_delete_data['material_impact'], extra_materials_file=extra_materials_json_path) # Pass extra_materials_file

        if delete_elevation_type in current_saved_elevations:
            del current_saved_elevations[delete_elevation_type]

        try:
            with open(elevations_json_path, 'w') as f: # Use elevations_json_path
                json.dump(current_saved_elevations, f, indent=4)
        except IOError as e:
            print(f"Error saving updated {elevations_json_path} during delete: {e}")
            if completion_callback: completion_callback(f"Error saving updated elevations after delete: {e}")
            return
        
    elif mode == "export_all":
        # In export_all mode, we don't modify current_saved_elevations here.
        # We just proceed to rebuild the Excel based on the existing JSON.
        pass
    else: # This is the "save_or_update" mode, or default behavior
        if elevation_type in current_saved_elevations and not reset: 
            old_elevation_data = current_saved_elevations[elevation_type]
            if 'material_impact' in old_elevation_data:
                reverse_material_impact(old_elevation_data['material_impact'], extra_materials_file=extra_materials_json_path) # Pass extra_materials_file
        
        current_saved_elevations[elevation_type] = {
            "system": system_input, "finish": finish_input, "total_count": total_count,
            "bays_wide": bays_wide, "bays_tall": bays_tall, "opening_width_inches": opening_width,
            "opening_height_inches": opening_height, "sqft_per_type": sqft_per_type, "total_sqft": total_sqft,
            "perimeter_ft": perimeter_ft, "total_perimeter_ft": total_perimeter_ft,
            "calculated_outputs": calculated_outputs, "material_impact": []
        }
        if door_size is not None: current_saved_elevations[elevation_type]['door_size'] = door_size

        try:
            with open(elevations_json_path, 'w') as f: # Use elevations_json_path
                json.dump(current_saved_elevations, f, indent=4)
        except IOError as e:
            print(f"Error saving elevation to {elevations_json_path}: {e}")
            if completion_callback: completion_callback(f"Error saving elevation: {e}")
            return

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    
    # Reset extra materials to an empty state at the beginning of a full report rebuild
    # Then load them to ensure we start calculations with a clean slate for impacts
    save_extra_materials({}, extra_materials_json_path) # Pass extra_materials_json_path
    overall_current_extra_materials_state = load_extra_materials(extra_materials_json_path) # Pass extra_materials_json_path

    current_excel_row = 1
    sorted_elev_names = sorted(current_saved_elevations.keys())

    if not sorted_elev_names:
        _clean_trailing_blank_rows(ws, 1)
    else:
        for elev_name in sorted_elev_names:
            elev_data = current_saved_elevations[elev_name]
            
            input_data = [
                ("System Input", elev_data.get("system")), ("Elevation Type", elev_name), ("Total Count", elev_data.get("total_count")),
                ("Bays Wide", elev_data.get("bays_wide")), ("Bays Tall", elev_data.get("bays_tall")), ("Opening Width", elev_data.get("opening_width_inches")),
                ("Opening Height", elev_data.get("opening_height_inches")), ("Sq Ft per Type", elev_data.get("sqft_per_type")), ("Total Sq Ft", elev_data.get("total_sqft")),
                ("Perimeter Ft", elev_data.get("perimeter_ft")), ("Total Perimeter Ft", elev_data.get("total_perimeter_ft"))
            ]
            if elev_data.get("door_size") is not None: input_data.append(("Door Size", elev_data.get("door_size")))

            for i, (header, value) in enumerate(input_data):
                ws.cell(row=current_excel_row + i, column=COL_A, value=header).font = Font(bold=True)
                ws.cell(row=current_excel_row + i, column=COL_B, value=value)

            output_section_current_row = current_excel_row 
            
            profiles_for_section, accessories_for_section, other_items_for_section = [], [], [] # Removed manual_pn_items_for_section

            for item in elev_data.get('calculated_outputs', []):
                pn, manual = item.get('part_number'), item.get('manual', False)
                if pn and pn != "N/A": # Item has a part number
                    if manual: # It's a manual item with a part number, group by its 'type'
                        other_items_for_section.append(item)
                    elif pn in PART_NUMBER_MAP.get("profiles", []):
                        profiles_for_section.append(item)
                    elif pn in PART_NUMBER_MAP.get("accessories", []):
                        accessories_for_section.append(item)
                    else: # Non-manual, with PN, but not profile/accessory
                        other_items_for_section.append(item)
                else: # Item does not have a valid part number (manual_no_pn or other misc)
                    other_items_for_section.append(item)
            
            multiplier = {"clear": 1.0, "black": 1.1, "paint": 1.2}.get(elev_data.get("finish").lower(), 1.0)
            system_total_for_this_block = [0.0]
            newly_calculated_material_impacts_for_this_elevation = []

            next_row_after_profiles, impacts_p = _write_output_section(ws, "PROFILES", profiles_for_section, COL_E, multiplier, system_total_for_this_block, output_section_current_row, overall_current_extra_materials_state, extra_materials_json_path) # Pass extra_materials_json_path
            next_row_after_accessories, impacts_a = _write_output_section(ws, "ACCESSORIES", accessories_for_section, COL_E, multiplier, system_total_for_this_block, next_row_after_profiles, overall_current_extra_materials_state, extra_materials_json_path) # Pass extra_materials_json_path
            # Removed the call for "MANUAL PART-NUMBERED ITEMS"
            
            newly_calculated_material_impacts_for_this_elevation.extend(impacts_p)
            newly_calculated_material_impacts_for_this_elevation.extend(impacts_a)
            # No longer extending impacts_mpn as that section is removed

            current_section_row = next_row_after_accessories # Adjusted starting row for subsequent sections
            grouped_other_misc = {}; 
            for item in other_items_for_section:
                grouped_other_misc.setdefault(item.get('type', 'MISCELLANEOUS ITEMS').upper(), []).append(item)
            
            for grp_title, grp_items in grouped_other_misc.items():
                next_row_after_group, impacts_g = _write_output_section(ws, grp_title, grp_items, COL_E, 1.0, system_total_for_this_block, current_section_row, overall_current_extra_materials_state, extra_materials_json_path) # Pass extra_materials_json_path
                newly_calculated_material_impacts_for_this_elevation.extend(impacts_g)
                current_section_row = next_row_after_group

            current_saved_elevations[elev_name]['material_impact'] = newly_calculated_material_impacts_for_this_elevation

            system_total_row = ws.max_row + 2 
            ws.cell(row=system_total_row, column=PRICE_COL, value="SYSTEM TOTAL").font = Font(bold=True)
            ws.cell(row=system_total_row + 1, column=PRICE_COL, value=system_total_for_this_block[0]).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            print(f"Rebuilt System Total for '{elev_name}': ${system_total_for_this_block[0]:.2f}")
            
            current_excel_row = system_total_row + 3

    save_extra_materials(overall_current_extra_materials_state, extra_materials_json_path) # Pass extra_materials_json_path

    _recalculate_running_grand_total(ws, PRICE_COL)
    # Changed the order: Clean trailing rows BEFORE autofitting to ensure max_row is accurate.
    _clean_trailing_blank_rows(ws, 1)
    _autofit_columns(ws, COL_A, PRICE_COL, 1, ws.max_row) # Autofit now runs on the final, cleaned sheet
    
    try:
        wb.save(excel_path) # Use excel_path
        print(f"Excel report '{excel_path}' fully rebuilt.")
    except Exception as save_err:
        print(f"Error saving Excel report during full rebuild: {save_err}")
        if completion_callback: completion_callback(f"Error saving report: {save_err}")
        return

    try:
        with open(elevations_json_path, 'w') as f: # Use elevations_json_path
            json.dump(current_saved_elevations, f, indent=4)
    except IOError as e:
        print(f"Error saving all elevations to {elevations_json_path} after rebuild: {e}")

    create_summary_sheet(excel_path=excel_path, elevations_json_path=elevations_json_path, extra_materials_json_path=extra_materials_json_path) # Pass all three paths
    if completion_callback: completion_callback()
