import os
import json
from openpyxl import Workbook
from openpyxl.styles import Font, numbers
from openpyxl.utils import get_column_letter

# Removed global constants for file paths, as they will now be passed as arguments
from utils.pricing import get_price_by_part, reverse_material_impact, load_extra_materials, save_extra_materials, apply_material_impact_to_extra_materials_in_memory, get_unit_price_by_part
from data.part_number import PART_NUMBER_MAP
from utils.formulas import calculate_door_info

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
    """Autofits columns in the worksheet."""
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

def _write_output_section(ws, title, items, colE, elevation_finish, system_total_ref, start_output_row, current_extra_materials_state, extra_materials_path):
    """Writes a section of calculated outputs to the worksheet."""
    if not items: return start_output_row, []

    current_row = start_output_row
    ws.cell(row=current_row, column=colE, value=title).font = Font(bold=True)
    for i, h in enumerate(["Description", "Part Number", "Quantity", "Price"]):
        ws.cell(row=current_row + 1, column=colE + i, value=h).font = Font(bold=True)
    current_row += 2

    section_material_impacts = []

    for item in items:
        qty_raw = item.get('quantity', 0)
        pn, manual = item.get('part_number'), item.get('manual', False)

        # Normalize quantity to always be an iterable (list) for processing individual cuts
        individual_quantities = qty_raw if isinstance(qty_raw, list) else [qty_raw]

        # Calculate display string for the Excel cell BEFORE looping
        if isinstance(qty_raw, list):
            if len(qty_raw) > 1 and all(x == qty_raw[0] for x in qty_raw):
                display_qty_string = f"{qty_raw[0]:.2f} x {len(qty_raw)}"
            else:
                display_qty_string = ", ".join([f"{q:.2f}" for q in qty_raw])
        else:
            display_qty_string = f"{qty_raw:.2f}"

        item_total_cost_for_display = 0.0 # Accumulate price for this line item

        for single_qty_for_calc in individual_quantities: # Loop through each individual quantity
            total_item_price_single_cut, unit_type, material_impact_details = 0.0, "pcs", None

            if manual:
                if pn and pn != "N/A":
                    # Pass the finish for manual items with part numbers
                    price_calculated, unit_calculated, material_impact_details = \
                        get_price_by_part(pn, single_qty_for_calc, finish=elevation_finish, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False, group=True)  
                    total_item_price_single_cut = (price_calculated if price_calculated is not None else item.get('price', 0.0) * single_qty_for_calc)
                    unit_type = unit_calculated or item.get('unit', 'pcs')
                else:
                    total_item_price_single_cut = item.get('price', 0.0) * single_qty_for_calc
                    unit_type = item.get('unit', 'pcs')
                    # For manual items without PN, ensure material_impact_details has the new display format
                    material_impact_details = {
                        'part_number': "N/A - Manual", 'requested_qty': single_qty_for_calc, 'purchased_qty_or_length': 0.0,
                        'leftover_generated_qty_or_length': 0.0, 'used_from_leftover_qty_or_length': 0.0,
                        'cost_incurred': total_item_price_single_cut, 'type_processed_as': 'manual_no_pn',
                        'finish': None # No finish for manual items without PN
                    }
            else:
                # Pass the finish for non-manual items
                total_price, unit_type, material_impact_details = \
                    get_price_by_part(pn, single_qty_for_calc, finish=elevation_finish, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False)
                total_item_price_single_cut = total_price or 0.0
                unit_type = unit_type or "pcs"
            
            # Accumulate the cost for the current item in the Excel row
            item_total_cost_for_display += total_item_price_single_cut

            if material_impact_details:
                # Format leftover_generated_qty_or_length for display (it will be a single float here)
                material_impact_details['leftover_generated_qty_or_length_display'] = f"{material_impact_details.get('leftover_generated_qty_or_length', 0.0):.2f}"
                
                section_material_impacts.append(material_impact_details)
                apply_material_impact_to_extra_materials_in_memory(current_extra_materials_state, material_impact_details)

        # Removed multiplier application as pricing is now dynamic based on finish
        
        system_total_ref[0] += item_total_cost_for_display # Add the accumulated cost to grand total

        ws.cell(row=current_row, column=colE, value=item.get('description', ''))
        ws.cell(row=current_row, column=colE + 1, value=pn or 'N/A')
        # Display the original (possibly list-formatted) quantity with the determined unit type
        ws.cell(row=current_row, column=colE + 2, value=f"{display_qty_string} {unit_type}")
        ws.cell(row=current_row, column=colE + 3, value=item_total_cost_for_display).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
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
    """
    Reads elevation data, aggregates quantities and prices by part number,
    and writes a clean summary section into the Excel file, including reusable material data.
    """
    import json
    from openpyxl import load_workbook
    from openpyxl.styles import Font, numbers

    # === Load Elevations ===
    try:
        with open(elevations_json_path, 'r') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"⚠️ Could not load elevations JSON: {e}")
        return

    # === Load Extra Materials ===
    try:
        extra_materials = load_extra_materials(extra_materials_json_path)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"⚠️ Could not load extra materials JSON: {e}")
        extra_materials = {}

    # === Load Excel ===
    try:
        wb = load_workbook(excel_path)
        ws = wb.active
        ws.title = "Report"
    except Exception as e:
        print(f"⚠️ Could not open workbook: {e}")
        return

    _delete_summary_section(ws)

    if not data:
        wb.save(excel_path)
        print("ℹ️ No data found, summary cleared if existed.")
        return

    # === Aggregate ===
    aggregated = {}

    for elev in data.values():
        elevation_finish = elev.get('finish') # Get finish for the current elevation
        for output in elev.get('calculated_outputs', []):
            part = output.get('part_number', '').strip()
            desc = output.get('description', '').strip()
            manual = output.get('manual', False)
            qty = output.get('quantity', 0)
            
            # Aggregate quantities: sum if it's a list, otherwise use as is
            qty_for_aggregation = sum(qty) if isinstance(qty, list) else qty

            # Determine if the part is a profile to include finish in the key
            # This is critical for distinguishing materials by finish in aggregation
            is_profile_part = part in PART_NUMBER_MAP.get('profiles', {})

            if manual:
                if part and part != "N/A":
                    # For manual items with a part number, also include finish if it's a profile
                    key = f"MANUAL_{part}-{elevation_finish.lower()}" if is_profile_part and elevation_finish else f"MANUAL_{part}"
                    display = f"{desc} ({part} - {elevation_finish})" if is_profile_part and elevation_finish else f"{desc} ({part})"
                else:
                    key = f"MANUAL_NO_PN_{desc}"
                    display = desc
            else:
                # For non-manual profiles, include finish in the key for distinct aggregation
                if is_profile_part and elevation_finish:
                    key = f"{part}-{elevation_finish.lower()}"
                    display = f"{part} ({elevation_finish})"
                else:
                    key = part
                    display = part

            if key not in aggregated:
                aggregated[key] = {
                    'quantity': 0.0, # Initialize as float
                    'description': desc,
                    'display': display,
                    'part_number': part,
                    'manual': manual,
                    'price': output.get('price', 0.0), # Storing the per-unit price or base price
                    'unit': output.get('unit', 'pcs'),
                    'finish': elevation_finish if is_profile_part or (manual and part and part != "N/A" and is_profile_part) else None # Store finish for profiles (manual or not)
                }

            try:
                # Add the aggregated quantity for this item
                aggregated[key]['quantity'] += float(qty_for_aggregation)
            except (TypeError, ValueError):
                pass

# === Build Summary Rows ===
    final_summary_data = []

    for key, item in aggregated.items():
        quantity_aggregated = item['quantity'] # This is the summed quantity
        manual = item['manual']
        part = item['part_number']
        display = item['display']
        original_unit_from_item = item['unit']
        item_finish = item.get('finish') # Get the stored finish for this aggregated item

        total_cost_for_item = 0.0
        calculated_unit_type = original_unit_from_item 
        reusable_qty_sum = 0.0
        reusable_pct = 0.0
        reusable_cost = 0.0
        
        # Calculate total_cost_for_item and calculated_unit_type
        if manual:
            if part and part != "N/A":
                # For manual items with part numbers, calculate price using get_price_by_part
                price_from_part, unit_type_from_pricing, _ = get_price_by_part(part, quantity_aggregated, finish=item_finish, extra_materials_file=extra_materials_json_path, summary=True, group=True)
                total_cost_for_item = price_from_part if price_from_part is not None else 0.0
                # Prioritize original_unit_from_item for manual part-numbered items
                calculated_unit_type = original_unit_from_item or unit_type_from_pricing
            else:
                # For manual items without a part number, use the provided price * quantity
                try:
                    price = float(item['price'])
                except (TypeError, ValueError):
                    price = 0.0
                try:
                    qty_float = float(quantity_aggregated)
                except (TypeError, ValueError):
                    qty_float = 0.0
                total_cost_for_item = price * qty_float
                calculated_unit_type = original_unit_from_item
        else:
            # Pass the item_finish to get_price_by_part for non-manual items
            total_price, unit_type_from_pricing, _ = get_price_by_part(part, quantity_aggregated, finish=item_finish, extra_materials_file=extra_materials_json_path, summary=True)
            total_cost_for_item = total_price if total_price is not None else 0.0
            calculated_unit_type = unit_type_from_pricing or original_unit_from_item

        reusable_qty_display_string = "N/A" # Default for display
        
        # This block now correctly handles both non-manual and manual items with part numbers
        if part and part != "N/A":
            # Construct the key for extra materials based on part number and finish (for profiles)
            extra_materials_key_for_reuse = part
            is_profile_part_for_reuse = part in PART_NUMBER_MAP.get('profiles', {})

            # If it's a profile (manual or not) and has a finish, append finish to key
            if is_profile_part_for_reuse and item_finish:
                extra_materials_key_for_reuse = f"{part}-{item_finish.lower()}"

            part_data = extra_materials.get(extra_materials_key_for_reuse, {})
            
            # --- MODIFICATION START ---
            # Prioritize length_pieces if it exists, regardless of whether it's a 'profile'
            if part_data.get("length_pieces"):
                reusable_qty_sum = sum(float(x) for x in part_data["length_pieces"] if isinstance(x, (int, float, str)))
                reuse_lengths_formatted = [f"{float(x):.2f}" for x in part_data["length_pieces"] if isinstance(x, (int, float, str))]
                reusable_qty_display_string = ", ".join(reuse_lengths_formatted)
            else:
                # Fallback to 'quantity' field if length_pieces is not present or empty
                reusable_qty_sum = part_data.get("quantity", 0.0)
                reusable_qty_display_string = f"{float(reusable_qty_sum):.2f}"
            # --- MODIFICATION END ---

            # Ensure reusable_qty_sum is float for calculations
            try:
                reusable_qty_sum = float(reusable_qty_sum)
            except (TypeError, ValueError):
                reusable_qty_sum = 0.0

            try:
                quantity_aggregated_f = float(quantity_aggregated)
            except (TypeError, ValueError):
                quantity_aggregated_f = 0.0

            if quantity_aggregated_f > 0:
                reusable_pct = (reusable_qty_sum / quantity_aggregated_f) * 100
            else:
                reusable_pct = 0.0

            # Pass the item_finish to get_unit_price_by_part
            unit_price_for_reuse, unit_type_for_reusable_calc = get_unit_price_by_part(part, finish=item_finish, extra_materials_file=extra_materials_json_path)
            print(unit_price_for_reuse)
            reusable_cost = reusable_qty_sum * unit_price_for_reuse if unit_price_for_reuse is not None else 0.0
            
            # Use the unit from get_unit_price_by_part for reusable quantity display if available
            # This ensures consistency between total and reusable units if they are linked by part number
            
            # --- MODIFICATION START (already present, just for context) ---
            if manual and part and part != "N/A":
                # For manual part-numbered items, prioritize the 'unit' from the aggregated item,
                # then fall back to the unit from get_unit_price_by_part, then the calculated_unit_type.
                # This ensures the original input unit for manual items is respected.
                calculated_unit_type = original_unit_from_item or unit_type_for_reusable_calc or calculated_unit_type
            else:
                calculated_unit_type = unit_type_for_reusable_calc or calculated_unit_type
            # --- MODIFICATION END ---
        
        final_summary_data.append((
            display,
            f"{quantity_aggregated:.2f} {calculated_unit_type}", # Format quantity for display with determined unit
            total_cost_for_item,
            f"{reusable_qty_display_string} {calculated_unit_type}" if part and part != "N/A" else "N/A", # Use formatted string and determined unit
            reusable_pct,
            reusable_cost,
            part
        ))

    # === Write to Sheet ===
    last_gt = _find_row_by_value(ws, 8, "RUNNING GRAND TOTAL", reverse=True)
    start_row = (last_gt + 3) if last_gt else ws.max_row + 2

    if not final_summary_data:
        wb.save(excel_path)
        print("ℹ️ Nothing to summarize.")
        return

    headers = [
        "Part Number / Description", "Total Quantity", "Total Price",
        "Reusable Material Quantity", "Reusable % of Total", "Reusable Material Cost"
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=start_row, column=col, value=header).font = Font(bold=True)

    for idx, (display, qty_disp, total_cost, reuse_qty_disp, reuse_pct, reuse_cost, part) in enumerate(final_summary_data, start=start_row + 1):
        ws.cell(row=idx, column=1, value=display)
        ws.cell(row=idx, column=2, value=qty_disp)
        ws.cell(row=idx, column=3, value=total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

        if part and part != "N/A": # This condition now includes manual items with part numbers
            ws.cell(row=idx, column=4, value=reuse_qty_disp)
            ws.cell(row=idx, column=5, value=f"{reuse_pct:.2f}%")
            ws.cell(row=idx, column=6, value=reuse_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        else:
            ws.cell(row=idx, column=4, value="N/A")
            ws.cell(row=idx, column=5, value="N/A").number_format = numbers.FORMAT_TEXT
            ws.cell(row=idx, column=6, value="").number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE


    # --- Add Reusable Grand Total ---
    reuse_total = sum(row[5] for row in final_summary_data) 
    rg_total_row = start_row + len(final_summary_data) + 1

    ws.cell(row=rg_total_row, column=5, value="REUSABLE GRAND TOTAL").font = Font(bold=True)
    ws.cell(row=rg_total_row, column=6, value=reuse_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    running_gt_row = _find_row_by_value(ws, 8, "RUNNING GRAND TOTAL", reverse=True)
    overall_gt = ws.cell(row=running_gt_row + 1, column=8).value if running_gt_row else 0.0

    try:
        overall_gt = float(str(overall_gt).strip("$")) if isinstance(overall_gt, str) else float(overall_gt)
    except ValueError:
        overall_gt = 0.0

    reuse_pct_of_gt = (reuse_total / overall_gt * 100) if overall_gt else 0.0

    ws.cell(row=rg_total_row + 1, column=5, value="REUSABLE % OF RUNNING GRAND TOTAL").font = Font(bold=True)
    ws.cell(row=rg_total_row + 1, column=6, value=f"{reuse_pct_of_gt:.2f}%")

    _autofit_columns(ws, 1, 6, start_row, rg_total_row + 1)
    _clean_trailing_blank_rows(ws, 1)

    wb.save(excel_path)
    print(f"✅ Summary updated: {excel_path}")

def generate_excel_report(
    excel_path, elevations_json_path, extra_materials_json_path,
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, reset=False, delete_elevation_type=None,
    doors=None, mode=None
):
    """Generates or updates an Excel report with detailed elevation inputs and calculated outputs."""
    COL_A, COL_B, COL_E, PRICE_COL = 1, 2, 5, 8

    current_saved_elevations = {}
    if os.path.exists(elevations_json_path):
        try:
            with open(elevations_json_path, 'r') as f:
                current_saved_elevations = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Error loading {elevations_json_path}: {e}. Starting with empty elevations in memory.")

    if delete_elevation_type:
        elevation_to_delete_data = current_saved_elevations.get(delete_elevation_type)
        if elevation_to_delete_data and 'material_impact' in elevation_to_delete_data:
            reverse_material_impact(elevation_to_delete_data['material_impact'], extra_materials_file=extra_materials_json_path)

        if delete_elevation_type in current_saved_elevations:
            del current_saved_elevations[delete_elevation_type]

        try:
            with open(elevations_json_path, 'w') as f:
                json.dump(current_saved_elevations, f, indent=4)
        except IOError as e:
            print(f"Error saving updated {elevations_json_path} during delete: {e}")
            if completion_callback: completion_callback(f"Error saving updated elevations after delete: {e}")
            return

    elif mode == "export_all":
        pass
    else:
        if elevation_type in current_saved_elevations and not reset:
            old_elevation_data = current_saved_elevations[elevation_type]
            if 'material_impact' in old_elevation_data:
                reverse_material_impact(old_elevation_data['material_impact'], extra_materials_file=extra_materials_json_path)

        # Include doors in calculated_outputs as manual items
        door_items = calculate_door_info(doors) if doors else []
        calculated_outputs.extend(door_items)

        current_saved_elevations[elevation_type] = {
            "system": system_input, "finish": finish_input, "total_count": total_count,
            "bays_wide": bays_wide, "bays_tall": bays_tall, "opening_width_inches": opening_width,
            "opening_height_inches": opening_height, "sqft_per_type": sqft_per_type, "total_sqft": total_sqft,
            "perimeter_ft": perimeter_ft, "total_perimeter_ft": total_perimeter_ft,
            "calculated_outputs": calculated_outputs,
            "material_impact": []
        }

        try:
            with open(elevations_json_path, 'w') as f:
                json.dump(current_saved_elevations, f, indent=4)
        except IOError as e:
            print(f"Error saving elevation to {elevations_json_path}: {e}")
            if completion_callback: completion_callback(f"Error saving elevation: {e}")
            return

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    save_extra_materials({}, extra_materials_json_path)
    overall_current_extra_materials_state = load_extra_materials(extra_materials_json_path)

    current_excel_row = 1
    sorted_elev_names = sorted(current_saved_elevations.keys())

    if not sorted_elev_names:
        _clean_trailing_blank_rows(ws, 1)
    else:
        for elev_name in sorted_elev_names:
            elev_data = current_saved_elevations[elev_name]

            input_data = [
                ("System Input", elev_data.get("system")),
                ("Finish", elev_data.get("finish")),
                ("Elevation Type", elev_name), ("Total Count", elev_data.get("total_count")),
                ("Bays Wide", elev_data.get("bays_wide")), ("Bays Tall", elev_data.get("bays_tall")),
                ("Opening Width", elev_data.get("opening_width_inches")),
                ("Opening Height", elev_data.get("opening_height_inches")),
                ("Sq Ft per Type", elev_data.get("sqft_per_type")),
                ("Total Sq Ft", elev_data.get("total_sqft")),
                ("Perimeter Ft", elev_data.get("perimeter_ft")),
                ("Total Perimeter Ft", elev_data.get("total_perimeter_ft"))
            ]

            for i, (header, value) in enumerate(input_data):
                ws.cell(row=current_excel_row + i, column=COL_A, value=header).font = Font(bold=True)
                ws.cell(row=current_excel_row + i, column=COL_B, value=value)

            output_section_current_row = current_excel_row

            profiles_for_section, accessories_for_section, other_items_for_section = [], [], []

            current_elevation_finish = elev_data.get("finish")

            for item in elev_data.get('calculated_outputs', []):
                pn, manual = item.get('part_number'), item.get('manual', False)
                if pn and pn != "N/A":
                    if manual:
                        other_items_for_section.append(item)
                    elif pn in PART_NUMBER_MAP.get("profiles", []):
                        profiles_for_section.append(item)
                    elif pn in PART_NUMBER_MAP.get("accessories", []):
                        accessories_for_section.append(item)
                    else:
                        other_items_for_section.append(item)
                else:
                    other_items_for_section.append(item)

            system_total_for_this_block = [0.0]
            newly_calculated_material_impacts_for_this_elevation = []

            next_row_after_profiles, impacts_p = _write_output_section(
                ws, "PROFILES", profiles_for_section, COL_E, current_elevation_finish,
                system_total_for_this_block, output_section_current_row,
                overall_current_extra_materials_state, extra_materials_json_path
            )

            next_row_after_accessories, impacts_a = _write_output_section(
                ws, "ACCESSORIES", accessories_for_section, COL_E, current_elevation_finish,
                system_total_for_this_block, next_row_after_profiles,
                overall_current_extra_materials_state, extra_materials_json_path
            )

            newly_calculated_material_impacts_for_this_elevation.extend(impacts_p)
            newly_calculated_material_impacts_for_this_elevation.extend(impacts_a)

            current_section_row = next_row_after_accessories
            grouped_other_misc = {}

            for item in other_items_for_section:
                item_type = item.get('type', 'MISCELLANEOUS ITEMS').upper()
                grouped_other_misc.setdefault(item_type, []).append(item)

            for grp_title, grp_items in grouped_other_misc.items():
                next_row_after_group, impacts_g = _write_output_section(
                    ws, grp_title, grp_items, COL_E, None,
                    system_total_for_this_block, current_section_row,
                    overall_current_extra_materials_state, extra_materials_json_path
                )
                newly_calculated_material_impacts_for_this_elevation.extend(impacts_g)
                current_section_row = next_row_after_group

            current_saved_elevations[elev_name]['material_impact'] = newly_calculated_material_impacts_for_this_elevation

            system_total_row = ws.max_row + 2
            ws.cell(row=system_total_row, column=PRICE_COL, value="SYSTEM TOTAL").font = Font(bold=True)
            ws.cell(row=system_total_row + 1, column=PRICE_COL, value=system_total_for_this_block[0]).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            print(f"Rebuilt System Total for '{elev_name}': ${system_total_for_this_block[0]:.2f}")

            current_excel_row = system_total_row + 3

    save_extra_materials(overall_current_extra_materials_state, extra_materials_json_path)

    _recalculate_running_grand_total(ws, PRICE_COL)
    _clean_trailing_blank_rows(ws, 1)
    _autofit_columns(ws, COL_A, PRICE_COL, 1, ws.max_row)

    try:
        wb.save(excel_path)
        print(f"Excel report '{excel_path}' fully rebuilt.")
    except Exception as save_err:
        print(f"❌ Error saving Excel report during full rebuild: {save_err}")
        if completion_callback: completion_callback(f"Error saving report: {save_err}")
        return

    try:
        with open(elevations_json_path, 'w') as f:
            json.dump(current_saved_elevations, f, indent=4)
    except IOError as e:
        print(f"Error saving all elevations to {elevations_json_path} after rebuild: {e}")

    create_summary_sheet(
        excel_path=excel_path,
        elevations_json_path=elevations_json_path,
        extra_materials_json_path=extra_materials_json_path
    )

    if completion_callback:
        completion_callback()
