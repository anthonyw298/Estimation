import os
import json
import math
from openpyxl import Workbook
from openpyxl.styles import Font, numbers, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import Counter
import datetime

from utils.pricing import get_price_by_part, reverse_material_impact, load_extra_materials, save_extra_materials, apply_material_impact_to_extra_materials_in_memory, get_unit_price_by_part, parse_length_to_feet
from data.part_number import PART_NUMBER_MAP
from data.parts_data import parts_data
from utils.formulas import calculate_door_info

# --- Helper Functions ---

def _get_multiplier(running_grand_total):
    """Returns multiplier based on running grand total."""
    return 0.584 if running_grand_total < 50000 else 0.522

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
    

def _write_output_section(ws, title, items, colE, elevation_finish, system_total_ref, original_system_total_ref, start_output_row, current_extra_materials_state, extra_materials_path, multiplier):
    """Writes a section of calculated outputs to the worksheet."""
    if not items: return start_output_row, []

    current_row = start_output_row
    title_cell = ws.cell(row=current_row, column=colE, value=title)
    title_cell.font = Font(bold=True, size=12)
    # title_cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # Removed color fill for professional look

    for i, h in enumerate(["Description", "Part Number", "Quantity", "Original Price", "Discounted Price"]):
        header_cell = ws.cell(row=current_row + 1, column=colE + i, value=h)
        header_cell.font = Font(bold=True)
        header_cell.border = Border(bottom=Side(style='thin'))
        # header_cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid") # Removed color fill for professional look
    current_row += 2

    section_material_impacts = []
    section_original_total = 0.0
    section_discounted_total = 0.0

    for item in items:
        qty_raw = item.get('quantity', 0)
        pn, manual = item.get('part_number'), item.get('manual', False)
        is_profile = pn in PART_NUMBER_MAP.get('profiles', {})
        is_accessory = pn in PART_NUMBER_MAP.get('accessories', {}) or item.get('type', '').lower() == 'accessory'
        is_glass = pn == "GLASS_AREA" or item.get('type', '').lower() == 'glass'

        individual_quantities = qty_raw if isinstance(qty_raw, list) else [qty_raw]
        qty_sum = sum(individual_quantities)

        unit_type = 'ft' if is_profile else 'pcs' if is_accessory else item.get('unit', 'pcs' if not is_glass else 'sqft')
        display_unit = unit_type

        if isinstance(qty_raw, list):
            if len(qty_raw) > 1 and all(x == qty_raw[0] for x in qty_raw):
                display_qty_string = f"{qty_raw[0]:.2f} {display_unit} x {len(qty_raw)}"
            else:
                display_qty_string = ", ".join([f"{q:.2f} {display_unit}" for q in qty_raw])
        else:
            display_qty_string = f"{qty_raw:.2f} {display_unit}"

        item_total_cost_for_display = 0.0
        original_item_total_cost = 0.0

        for single_qty_for_calc in individual_quantities:
            total_item_price_single_cut, calculated_unit_type, material_impact_details = 0.0, unit_type, None

            if manual:
                if pn and pn != "N/A":
                    price_calculated, unit_calculated, material_impact_details = \
                        get_price_by_part(pn, single_qty_for_calc, finish=elevation_finish, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False, group=True)
                    total_item_price_single_cut = (price_calculated if price_calculated is not None else item.get('price', 0.0) * single_qty_for_calc)
                    calculated_unit_type = unit_type if is_profile or is_accessory else (unit_calculated or item.get('unit', 'pcs'))
                else:
                    total_item_price_single_cut = item.get('price', 0.0) * single_qty_for_calc
                    calculated_unit_type = item.get('unit', 'pcs')
                    material_impact_details = {
                        'part_number': "N/A - Manual", 'requested_qty': single_qty_for_calc, 'purchased_qty_or_length': 0.0,
                        'leftover_generated_qty_or_length': 0.0, 'used_from_leftover_qty_or_length': 0.0,
                        'cost_incurred': total_item_price_single_cut, 'type_processed_as': 'manual_no_pn',
                        'finish': None
                    }
            else:
                total_price, unit_from_pricing, material_impact_details = \
                    get_price_by_part(pn, single_qty_for_calc, finish=elevation_finish, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False)
                total_item_price_single_cut = total_price or 0.0
                calculated_unit_type = unit_type if is_profile or is_accessory else (unit_from_pricing or item.get('unit', 'pcs'))

            item_total_cost_for_display += total_item_price_single_cut
            original_item_total_cost += total_item_price_single_cut

            if material_impact_details:
                leftover_qty = material_impact_details.get('leftover_generated_qty_or_length', 0.0)
                material_impact_details['leftover_generated_qty_or_length_display'] = f"{leftover_qty:.2f} {display_unit}"
                section_material_impacts.append(material_impact_details)
                apply_material_impact_to_extra_materials_in_memory(current_extra_materials_state, material_impact_details)

        if is_profile or is_accessory:
            item_total_cost_for_display *= multiplier
            if qty_sum > 0:
                item['price'] = item_total_cost_for_display / qty_sum

        system_total_ref[0] += item_total_cost_for_display
        original_system_total_ref[0] += original_item_total_cost
        section_original_total += original_item_total_cost
        section_discounted_total += item_total_cost_for_display

        ws.cell(row=current_row, column=colE, value=item.get('description', ''))
        ws.cell(row=current_row, column=colE + 1, value=pn or 'N/A')
        ws.cell(row=current_row, column=colE + 2, value=display_qty_string)
        ws.cell(row=current_row, column=colE + 3, value=original_item_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=colE + 4, value=item_total_cost_for_display).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        current_row += 1

    # Add Section Totals
    ws.cell(row=current_row, column=colE + 2, value=f"Total {title}").font = Font(bold=True)
    ws.cell(row=current_row, column=colE + 3, value=section_original_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=current_row, column=colE + 3).font = Font(bold=True)
    ws.cell(row=current_row, column=colE + 4, value=section_discounted_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=current_row, column=colE + 4).font = Font(bold=True)
    
    # Add top border for totals row
    for col in range(colE, colE + 5):
        ws.cell(row=current_row, column=col).border = Border(top=Side(style='thin'))

    return current_row + 2, section_material_impacts


def create_summary_sheet(ws, elevations_json_path, extra_materials_json_path):
    """
    Reads elevation data, aggregates quantities and prices by part number across all elevations,
    and writes a clean summary section into the worksheet, grouped by profiles, accessories, doors, glass, and labor.
    """
    try:
        with open(elevations_json_path, 'r') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"⚠️ Could not load elevations JSON: {e}")
        return

    try:
        extra_materials = load_extra_materials(extra_materials_json_path)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"⚠️ Could not load extra materials JSON: {e}")
        extra_materials = {}

    if not data:
        print("ℹ️ No data found, summary cleared if existed.")
        return

    # Step 1: Calculate full_running_grand_total for multiplier
    full_running_grand_total = 0.0
    for elev_key, elev in data.items():
        elevation_finish = elev.get('finish', '').lower()
        for output in elev.get('calculated_outputs', []):
            qty = output.get('quantity', 0)
            qty = sum(qty) if isinstance(qty, list) else qty
            manual = output.get('manual', False)
            part = output.get('part_number', '').strip()
            item_type = output.get('type', '').lower()
            price = 0.0
            if manual or part == "GLASS_AREA" or item_type in ['glass', 'joints_fab_labor', 'door', 'doors']:
                price = output.get('price', 0.0) * qty
            else:
                price, _, _ = get_price_by_part(
                    part,
                    qty,
                    finish=elevation_finish,
                    extra_materials_file=extra_materials_json_path,
                    summary=True,
                    group=True if manual else False
                )
                price = price if price is not None else 0.0
            full_running_grand_total += price

    multiplier = _get_multiplier(full_running_grand_total)

    # Step 2: Aggregate quantities and prices across all elevations, grouped by category
    categories = {
        'PROFILES': [],
        'ACCESSORIES': [],
        'DOORS': [],
        'GLASS': [],
        'LABOR': []
    }

    for elev_key, elev in data.items():
        elevation_finish = elev.get('finish', '').lower()
        for output in elev.get('calculated_outputs', []):
            part = output.get('part_number', '').strip()
            desc = output.get('description', '').strip()
            manual = output.get('manual', False)
            qty = output.get('quantity', 0)
            qty_for_aggregation = sum(qty) if isinstance(qty, list) else qty
            is_profile = part in PART_NUMBER_MAP.get('profiles', {})
            is_accessory = part in PART_NUMBER_MAP.get('accessories', {}) or output.get('type', '').lower() == 'accessory'
            is_glass = part == "GLASS_AREA" or output.get('type', '').lower() == 'glass'
            is_joints_fab_labor = part == "JOINTS_FAB_LABOR" or output.get('type', '').lower() == 'joints_fab_labor' or "joints fabrication" in desc.lower() or "fabrication labor" in desc.lower()
            is_door = output.get('type', '').lower() in ['door', 'doors']

            if is_profile:
                category = 'PROFILES'
            elif is_accessory:
                category = 'ACCESSORIES'
            elif is_door:
                category = 'DOORS'
            elif is_glass:
                category = 'GLASS'
            elif is_joints_fab_labor:
                category = 'LABOR'
            else:
                continue

            if manual or is_glass or is_joints_fab_labor or is_door:
                if part and part != "N/A":
                    key = f"MANUAL_{part}-{elevation_finish}" if (is_profile or is_joints_fab_labor or is_door or is_glass) and elevation_finish else f"MANUAL_{part}"
                    display = f"{desc} ({part} - {elevation_finish})" if (is_profile or is_joints_fab_labor or is_door or is_glass) and elevation_finish else f"{desc} ({part})"
                else:
                    key = f"MANUAL_NO_PN_{desc}"
                    display = desc
            else:
                if (is_profile or is_joints_fab_labor or is_door or is_glass) and elevation_finish:
                    key = f"{part}-{elevation_finish}"
                    display = f"{part} ({elevation_finish})"
                else:
                    key = part
                    display = part

            categories[category].append({
                'key': key,
                'quantity': qty_for_aggregation,
                'quantity_list': qty if isinstance(qty, list) else [qty],  # Preserve individual quantities
                'description': desc,
                'display': display,
                'part_number': part,
                'manual': manual,
                'unit': 'ft' if is_profile else 'pcs' if is_accessory else output.get('unit', 'pcs' if not is_glass else 'sqft'),
                'finish': elevation_finish if (is_profile or is_joints_fab_labor or is_door or is_glass) else '',
                'is_glass': is_glass,
                'is_joints_fab_labor': is_joints_fab_labor,
                'is_door': is_door,
                'price': output.get('price', 0.0) if (manual or is_glass or is_joints_fab_labor or is_door) else 0.0
            })

    # Step 2.5: Aggregate items within each category by key to prevent duplicates across elevations
    for category in categories:
        aggregated_map = {}
        for item in categories[category]:
            k = item['key']
            if k in aggregated_map:
                existing = aggregated_map[k]
                # If price is manually set (manual, glass, labor, door), maintain correct total cost by updating unit price
                if existing['manual'] or existing['is_glass'] or existing['is_joints_fab_labor'] or existing['is_door']:
                    # Recalculate weighted average price
                    cost_existing = float(existing.get('price', 0.0)) * float(existing['quantity'])
                    cost_new = float(item.get('price', 0.0)) * float(item['quantity'])
                    total_qty = float(existing['quantity']) + float(item['quantity'])
                    existing['quantity'] = total_qty
                    # Combine quantity lists
                    existing_qty_list = existing.get('quantity_list', [])
                    item_qty_list = item.get('quantity_list', [])
                    if not existing_qty_list:
                        existing_qty_list = [existing['quantity']]
                    if not item_qty_list:
                        item_qty_list = [item['quantity']]
                    existing['quantity_list'] = existing_qty_list + item_qty_list
                    if total_qty > 0:
                        existing['price'] = (cost_existing + cost_new) / total_qty
                else:
                    # For standard parts, just sum quantity; price is re-calculated in Step 3
                    existing['quantity'] = float(existing['quantity']) + float(item['quantity'])
                    # Combine quantity lists
                    existing_qty_list = existing.get('quantity_list', [])
                    item_qty_list = item.get('quantity_list', [])
                    if not existing_qty_list:
                        existing_qty_list = [existing['quantity']]
                    if not item_qty_list:
                        item_qty_list = [item['quantity']]
                    existing['quantity_list'] = existing_qty_list + item_qty_list
            else:
                # Ensure quantity is float
                item['quantity'] = float(item['quantity'])
                if 'quantity_list' not in item:
                    item['quantity_list'] = [item['quantity']]
                aggregated_map[k] = item
        categories[category] = list(aggregated_map.values())

    # Step 3: Calculate prices for aggregated items and prepare final data
    final_summary_data = []
    total_discounted_price = 0.0
    total_reusable_cost = 0.0
    grand_original_total = 0.0
    grand_discounted_total = 0.0
    grand_residual_total = 0.0

    for category, items in categories.items():
        for item in items:
            key = item['key']
            quantity_aggregated = item['quantity']
            manual = item['manual']
            part = item['part_number']
            display = item['display']
            is_profile = part in PART_NUMBER_MAP.get('profiles', {})
            is_accessory = part in PART_NUMBER_MAP.get('accessories', {}) or item.get('type', '').lower() == 'accessory'
            is_glass = item['is_glass']
            is_joints_fab_labor = item['is_joints_fab_labor']
            is_door = item['is_door']
            item_finish = item['finish']

            display_unit = 'ft' if is_profile else 'pcs' if is_accessory else item['unit']
            original_total_cost_for_item = 0.0
            total_cost_for_item = 0.0
            calculated_unit_type = display_unit
            reusable_qty_sum = 0.0
            reusable_pct = 0.0
            reusable_cost = 0.0
            reusable_qty_display_string = "N/A"

            if manual or is_glass or is_joints_fab_labor or is_door:
                price = float(item.get('price', 0.0))
                qty_float = float(quantity_aggregated)
                original_total_cost_for_item = price * qty_float
                calculated_unit_type = item['unit'] or ('sqft' if is_glass else 'pcs')
            else:
                total_price, unit_type_from_pricing, _ = get_price_by_part(
                    part,
                    quantity_aggregated,
                    finish=item_finish,
                    extra_materials_file=extra_materials_json_path,
                    summary=True
                )
                original_total_cost_for_item = total_price if total_price is not None else 0.0
                calculated_unit_type = 'ft' if is_profile else 'pcs' if is_accessory else (unit_type_from_pricing or item['unit'] or 'pcs')

            if is_profile or is_accessory:
                total_cost_for_item = original_total_cost_for_item * multiplier
            else:
                total_cost_for_item = original_total_cost_for_item

            total_discounted_price += total_cost_for_item

            if part and part != "N/A" and (is_profile or is_accessory):
                extra_materials_key_for_reuse = part
                if is_profile and item_finish:
                    extra_materials_key_for_reuse = f"{part}-{item_finish}"

                part_data = extra_materials.get(extra_materials_key_for_reuse, {})
                if part_data.get("length_pieces"):
                    lengths = [float(x) for x in part_data["length_pieces"] if isinstance(x, (int, float, str))]
                    reusable_qty_sum = sum(lengths)
                    if lengths:
                        counter = Counter([f"{l:.2f}" for l in lengths])
                        reuse_lengths_formatted = [f"{length} {display_unit} x{count}" if count > 1 else f"{length} {display_unit}" for length, count in sorted(counter.items(), key=lambda x: float(x[0]))]
                        reusable_qty_display_string = ", ".join(reuse_lengths_formatted)
                else:
                    reusable_qty_sum = part_data.get("quantity", 0.0)
                    reusable_qty_display_string = f"{float(reusable_qty_sum):.2f} {display_unit}"

                try:
                    reusable_qty_sum = float(reusable_qty_sum)
                except (TypeError, ValueError):
                    reusable_qty_sum = 0.0

                try:
                    quantity_aggregated_f = float(quantity_aggregated)
                except (TypeError, ValueError):
                    quantity_aggregated_f = 0.0

                if reusable_qty_sum > 0 and quantity_aggregated_f > 0:
                    reusable_pct = min((reusable_qty_sum / (quantity_aggregated_f + reusable_qty_sum)) * 100, 100.0)
                else:
                    reusable_pct = 0.0

                unit_price_for_reuse, unit_type_for_reusable_calc = get_unit_price_by_part(
                    part, finish=item_finish, extra_materials_file=extra_materials_json_path
                )
                reusable_cost = (
                    reusable_qty_sum * unit_price_for_reuse * multiplier
                    if unit_price_for_reuse is not None
                    else 0.0
                )
                total_reusable_cost += reusable_cost
            # Calculate Quantity Req (FT) and Qty Stick (Req)
            quantity_req_ft = "N/A"
            qty_stick_req = "N/A"
            quantity_display_formatted = f"{quantity_aggregated:.2f} {display_unit}"
            
            if is_profile and part and part != "N/A":
                # For profiles, quantity is in feet - add units
                quantity_req_ft = f"{quantity_aggregated:.2f} ft"
                # Calculate number of sticks needed and format with stick length
                part_data = parts_data.get(part, {})
                length_str = part_data.get('Length', '')
                min_purchase_length = parse_length_to_feet(length_str) or 1.0
                if min_purchase_length > 0:
                    num_sticks = math.ceil(quantity_aggregated / min_purchase_length)
                    qty_stick_req = f"{num_sticks} ({min_purchase_length:.0f}ft per)"
                else:
                    qty_stick_req = "N/A"
                
                # Format quantity_display to show breakdown like "16ft x2, 8ft x1"
                quantity_list = item.get('quantity_list', [quantity_aggregated])
                # Count occurrences of each length
                length_counter = Counter([round(q, 2) for q in quantity_list])
                if len(length_counter) > 1:
                    # Multiple different lengths - show breakdown
                    length_parts = []
                    for length_val, count in sorted(length_counter.items(), key=lambda x: x[0], reverse=True):
                        if count > 1:
                            length_parts.append(f"{length_val:.0f}ft x{count}")
                        else:
                            length_parts.append(f"{length_val:.0f}ft x1")
                    quantity_display_formatted = ", ".join(length_parts)
                else:
                    # All same length - show total with count if multiple pieces
                    length_val = list(length_counter.keys())[0]
                    count = length_counter[length_val]
                    if count > 1:
                        quantity_display_formatted = f"{length_val:.0f}ft x{count}"
                    else:
                        quantity_display_formatted = f"{length_val:.0f}ft"
                    
            elif is_accessory and part and part != "N/A":
                # For accessories, get bulk order info
                part_data = parts_data.get(part, {})
                units_str = part_data.get('Units', '1 pcs.')
                unit_count_per_bundle = 1
                if 'pc' in units_str.lower():
                    try:
                        unit_count_per_bundle = int(units_str.lower().split('pc')[0].strip()) or 1
                    except ValueError:
                        unit_count_per_bundle = 1
                
                quantity_req_ft = f"{unit_count_per_bundle} pcs"
                num_orders = math.ceil(quantity_aggregated / unit_count_per_bundle) if unit_count_per_bundle > 0 else 0
                qty_stick_req = f"{num_orders} order{'s' if num_orders != 1 else ''}"
            else:
                # For other items (glass, labor, doors), show N/A in first column, unit price in second
                quantity_req_ft = "N/A"
                
                # For the second column, show unit price
                if quantity_aggregated > 0:
                    unit_price = original_total_cost_for_item / quantity_aggregated
                    qty_stick_req = f"${unit_price:.2f}"
                else:
                    qty_stick_req = "$0.00"
            
            final_summary_data.append({
                'category': category,
                'display': display,
                'quantity_display': quantity_display_formatted,
                'quantity_req_ft': quantity_req_ft,
                'qty_stick_req': qty_stick_req,
                'original_total_cost': original_total_cost_for_item,
                'total_cost': total_cost_for_item,
                'reusable_qty_display': reusable_qty_display_string,
                'reusable_pct': reusable_pct if (is_profile or is_accessory) else "N/A",
                'reusable_cost': reusable_cost if (is_profile or is_accessory) else 0.0,
                'part': part,
                'calculated_unit_type': calculated_unit_type
            })

    # Step 4: Write to worksheet with grouped sections
    start_row = 1
    current_row = start_row
    
    # Define headers based on category
    def get_headers_for_category(category, items_list=None):
        if category == 'PROFILES':
            return [
                "Project Total Materials", "Total Ft", "Sticks Req", "Quantity", "List Price", "Discounted List Price",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'ACCESSORIES':
            return [
                "Project Total Materials", "Quantity Per Order", "Orders Req", "Quantity", "List Price", "Discounted List Price",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'GLASS':
            return [
                "Project Total Materials", "N/A", "Unit Price", "Quantity", "List Price", "Discounted List Price",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'LABOR':
            return [
                "Project Total Materials", "N/A", "Unit Price", "Quantity", "List Price", "Discounted List Price",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'DOORS':
            return [
                "Project Total Materials", "N/A", "Unit Price", "Quantity", "List Price", "Discounted List Price",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        else:
            # Default headers
            return [
                "Project Total Materials", "Quantity Req (FT)", "Qty Stick (Req)", "Quantity", "List Price", "Discounted List Price",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]

    for category, items in categories.items():
        if not items:
            continue
        headers = get_headers_for_category(category, items)
        header_cell = ws.cell(row=current_row, column=1, value=category)
        header_cell.font = Font(bold=True, size=12)
        # header_cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # Removed color fill for professional look
        current_row += 1
        for col, header in enumerate(headers, start=1):
            header_cell = ws.cell(row=current_row, column=col, value=header)
            header_cell.font = Font(bold=True)
            header_cell.border = Border(bottom=Side(style='thin'))
            # header_cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid") # Removed color fill for professional look
        current_row += 1
        
        section_original_total = 0.0
        section_total_cost = 0.0
        section_residual_total = 0.0
        
        for item in final_summary_data:
            if item['category'] == category:
                section_original_total += item['original_total_cost']
                section_total_cost += item['total_cost']
                section_residual_total += item['reusable_cost']
                
                ws.cell(row=current_row, column=1, value=item['display'])
                ws.cell(row=current_row, column=2, value=item['quantity_req_ft'])
                ws.cell(row=current_row, column=3, value=item['qty_stick_req'])
                ws.cell(row=current_row, column=4, value=item['quantity_display'])
                ws.cell(row=current_row, column=5, value=item['original_total_cost']).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                ws.cell(row=current_row, column=6, value=item['total_cost']).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                ws.cell(row=current_row, column=7, value=item['reusable_qty_display'])
                ws.cell(row=current_row, column=8, value=f"{item['reusable_pct']:.2f}%" if isinstance(item['reusable_pct'], (int, float)) else item['reusable_pct'])
                ws.cell(row=current_row, column=9, value=item['reusable_cost']).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                current_row += 1
        
        grand_original_total += section_original_total
        grand_discounted_total += section_total_cost
        grand_residual_total += section_residual_total

        # Add Section Totals for Summary
        ws.cell(row=current_row, column=4, value=f"Total {category}").font = Font(bold=True)
        ws.cell(row=current_row, column=5, value=section_original_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=5).font = Font(bold=True)
        ws.cell(row=current_row, column=6, value=section_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=6).font = Font(bold=True)
        ws.cell(row=current_row, column=9, value=section_residual_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=9).font = Font(bold=True)
        
        # Add top border for totals row
        for col in range(1, 10):
            ws.cell(row=current_row, column=col).border = Border(top=Side(style='thin'))

        current_row += 2

    # Grand Totals Block
    gt_row = current_row + 2
    
    # Original Total
    ws.cell(row=gt_row, column=5, value="Overall Total Price (List)").font = Font(bold=True)
    ws.cell(row=gt_row, column=5).alignment = Alignment(horizontal='right')
    ws.cell(row=gt_row, column=5).border = Border(left=Side(style='thin'), top=Side(style='thin'))
    
    ws.cell(row=gt_row, column=6, value=grand_original_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=gt_row, column=6).font = Font(bold=True)
    ws.cell(row=gt_row, column=6).border = Border(right=Side(style='thin'), top=Side(style='thin'))

    # Discounted Total
    ws.cell(row=gt_row+1, column=5, value="Overall Discounted Total").font = Font(bold=True)
    ws.cell(row=gt_row+1, column=5).alignment = Alignment(horizontal='right')
    ws.cell(row=gt_row+1, column=5).border = Border(left=Side(style='thin'))

    ws.cell(row=gt_row+1, column=6, value=grand_discounted_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=gt_row+1, column=6).font = Font(bold=True)
    ws.cell(row=gt_row+1, column=6).border = Border(right=Side(style='thin'))

    # Residual Cost
    reuse_total = total_reusable_cost
    ws.cell(row=gt_row+2, column=5, value="Overall Residual Cost").font = Font(bold=True)
    ws.cell(row=gt_row+2, column=5).alignment = Alignment(horizontal='right')
    ws.cell(row=gt_row+2, column=5).border = Border(left=Side(style='thin'))

    ws.cell(row=gt_row+2, column=6, value=reuse_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=gt_row+2, column=6).font = Font(bold=True)
    ws.cell(row=gt_row+2, column=6).border = Border(right=Side(style='thin'))

    # Waste %
    reuse_pct_of_gt = min((total_reusable_cost / total_discounted_price * 100) if total_discounted_price > 0 else 0.0, 100.0)
    ws.cell(row=gt_row+3, column=5, value="Overall Waste %").font = Font(bold=True)
    ws.cell(row=gt_row+3, column=5).alignment = Alignment(horizontal='right')
    ws.cell(row=gt_row+3, column=5).border = Border(left=Side(style='thin'), bottom=Side(style='thin'))

    ws.cell(row=gt_row+3, column=6, value=f"{reuse_pct_of_gt:.2f}%").font = Font(bold=True)
    ws.cell(row=gt_row+3, column=6).border = Border(right=Side(style='thin'), bottom=Side(style='thin'))

    _autofit_columns(ws, 1, 9, start_row, gt_row + 4)
    _clean_trailing_blank_rows(ws, 1)

    print(f"✅ Summary updated with grouped sections: Profiles, Accessories, Doors, Glass, Labor.")

def _format_door_summary(calculated_outputs):
    if not calculated_outputs:
        return ""
    door_lines = []
    for item in calculated_outputs:
        if item.get("type", "").lower() in ["door", "doors"] and item.get('manual', False):
            quantity = item.get("quantity", 1)
            style = item.get("Style", "").strip()
            price = item.get("price", 0.0)
            hardware = item.get("hardware", {})
            if not (isinstance(quantity, (int, float)) and quantity > 0):
                continue
            if not style or style.lower() == "unknown":
                continue
            enabled_hw = [hw for hw, enabled in hardware.items() if enabled]
            if enabled_hw:
                door_lines.append(f"{quantity} x {style} Door (${price:,.2f})\n  with: {', '.join(enabled_hw)}")
            else:
                door_lines.append(f"{quantity} x {style} Door (${price:,.2f})")
    return "; \n".join(door_lines) if door_lines else "None"

def generate_excel_report(
    excel_path, elevations_json_path, extra_materials_json_path,
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, reset=False, delete_elevation_type=None,
    doors=None, mode=None, custom_bay_widths=None, custom_bay_heights=None
):
    """Generates or updates an Excel report with detailed elevation inputs and calculated outputs."""
    COL_A, COL_B, COL_E, PRICE_COL = 1, 2, 5, 9

    project_root = os.getcwd()
    private_projects_dir = os.path.join(project_root, '.files')
    public_reports_dir = os.path.join(project_root, 'reports')

    os.makedirs(private_projects_dir, exist_ok=True)
    os.makedirs(public_reports_dir, exist_ok=True)
    
    private_elevations_path = os.path.join(private_projects_dir, os.path.basename(elevations_json_path))
    private_extra_materials_path = os.path.join(private_projects_dir, os.path.basename(extra_materials_json_path))
    private_excel_path = os.path.join(private_projects_dir, os.path.basename(excel_path))
    
    output_excel_path = public_reports_dir if mode == "export_all" else private_excel_path
    
    current_saved_elevations = {}
    if os.path.exists(private_elevations_path):
        try:
            with open(private_elevations_path, 'r') as f:
                current_saved_elevations = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Error loading {private_elevations_path}: {e}. Starting with empty elevations in memory.")

    if delete_elevation_type:
        elevation_to_delete_data = current_saved_elevations.get(delete_elevation_type)
        if elevation_to_delete_data and 'material_impact' in elevation_to_delete_data:
            reverse_material_impact(elevation_to_delete_data['material_impact'], extra_materials_file=private_extra_materials_path)

        if delete_elevation_type in current_saved_elevations:
            del current_saved_elevations[delete_elevation_type]

        try:
            with open(private_elevations_path, 'w') as f:
                json.dump(current_saved_elevations, f, indent=4)
        except IOError as e:
            print(f"Error saving updated {private_elevations_path} during delete: {e}")
            if completion_callback: completion_callback(f"Error saving updated elevations after delete: {e}")

    # Handle regeneration/cleanup case
    if delete_elevation_type and system_input == "": 
        pass # Don't add a new empty elevation if we are just deleting
    elif mode == "export_all":
        pass
    else:
        if elevation_type in current_saved_elevations and not reset:
            old_elevation_data = current_saved_elevations[elevation_type]
            if 'material_impact' in old_elevation_data:
                reverse_material_impact(old_elevation_data['material_impact'], extra_materials_file=private_extra_materials_path)

        # Build door items from UI doors, but avoid double-adding if base outputs already include doors
        base_outputs = list(calculated_outputs or [])
        
        # Remove any existing door entries from base_outputs to prevent duplication/stale data
        base_outputs = [item for item in base_outputs if not (item.get('type', '').lower() in ['door', 'doors'] and item.get('manual', False))]
        
        # Recalculate door items fresh from current inputs
        door_items = calculate_door_info(doors, finish_input) if doors else []
        
        # Avoid mutating the incoming list reference
        elevation_outputs = base_outputs + door_items

        current_saved_elevations[elevation_type] = {
            "system": system_input, "finish": finish_input, "total_count": total_count,
            "bays_wide": bays_wide, "bays_tall": bays_tall, "opening_width_inches": opening_width,
            "opening_height_inches": opening_height, "sqft_per_type": sqft_per_type, "total_sqft": total_sqft,
            "perimeter_ft": perimeter_ft, "total_perimeter_ft": total_perimeter_ft,
            "calculated_outputs": elevation_outputs,
            "material_impact": [],
            "custom_bay_widths": custom_bay_widths or [],
            "custom_bay_heights": custom_bay_heights or []
        }

        try:
            with open(private_elevations_path, 'w') as f:
                json.dump(current_saved_elevations, f, indent=4)
        except IOError as e:
            print(f"Error saving elevation to {private_elevations_path}: {e}")
            if completion_callback: completion_callback(f"Error saving elevation: {e}")
            return

    wb = Workbook()
    # Remove the default "Sheet" immediately so it doesn't end up in the final report
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    save_extra_materials({}, private_extra_materials_path)
    overall_current_extra_materials_state = load_extra_materials(private_extra_materials_path)

    full_running_grand_total = 0.0
    for elev_name in current_saved_elevations:
        elev_data = current_saved_elevations[elev_name]
        for item in elev_data.get('calculated_outputs', []):
            pn = item.get('part_number')
            qty_raw = item.get('quantity', 0)
            individual_quantities = qty_raw if isinstance(qty_raw, list) else [qty_raw]
            manual = item.get('manual', False)
            price_sum = 0.0
            for single_qty in individual_quantities:
                if manual:
                    if pn and pn != "N/A":
                        p, _, _ = get_price_by_part(pn, single_qty, finish=elev_data.get('finish'), summary=True, group=True)
                        price_sum += p if p is not None else item.get('price', 0.0) * single_qty
                    else:
                        price_sum += item.get('price', 0.0) * single_qty
                else:
                    p, _, _ = get_price_by_part(pn, single_qty, finish=elev_data.get('finish'), summary=True)
                    price_sum += p if p is not None else 0.0
            full_running_grand_total += price_sum

    multiplier = _get_multiplier(full_running_grand_total)

    sorted_elev_names = sorted(current_saved_elevations.keys())

    if not sorted_elev_names:
        pass
    else:
        for elev_name in sorted_elev_names:
            ws = wb.create_sheet(title=elev_name)
            elev_data = current_saved_elevations[elev_name]

            # Format custom bay dimensions for display
            custom_bay_widths_str = ", ".join([f"{w:.2f} in" for w in elev_data.get('custom_bay_widths', [])]) if elev_data.get('custom_bay_widths') else "Equal distribution"
            custom_bay_heights_str = ", ".join([f"{h:.2f} in" for h in elev_data.get('custom_bay_heights', [])]) if elev_data.get('custom_bay_heights') else "Equal distribution"

            input_data = [
                ("System Input", elev_data.get("system")),
                ("Finish", elev_data.get("finish")),
                ("Elevation Type", elev_name),
                ("Total Count", elev_data.get("total_count")),
                ("Bays Wide", elev_data.get("bays_wide")),
                ("Bays Tall", elev_data.get("bays_tall")),
                ("Custom Bay Widths", custom_bay_widths_str),
                ("Custom Bay Heights", custom_bay_heights_str),
                ("Opening Width", f"{elev_data.get('opening_width_inches'):.2f} in"),
                ("Opening Height", f"{elev_data.get('opening_height_inches'):.2f} in"),
                ("Sq Ft per Type", f"{elev_data.get('sqft_per_type'):.2f} sqft"),
                ("Total Sq Ft", f"{elev_data.get('total_sqft'):.2f} sqft"),
                ("Perimeter Ft", f"{elev_data.get('perimeter_ft'):.2f} ft"),
                ("Total Perimeter Ft", f"{elev_data.get('total_perimeter_ft'):.2f} ft"),
                ("Doors", _format_door_summary(elev_data.get("calculated_outputs", [])))
            ]

            current_excel_row = 1
            thin_border = Border(left=Side(style='thin'), 
                                 right=Side(style='thin'), 
                                 top=Side(style='thin'), 
                                 bottom=Side(style='thin'))

            for i, (header, value) in enumerate(input_data):
                header_cell = ws.cell(row=current_excel_row + i, column=COL_A, value=header)
                header_cell.font = Font(bold=True)
                # header_cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # Removed color fill for professional look
                header_cell.border = thin_border 

                value_cell = ws.cell(row=current_excel_row + i, column=COL_B, value=value)
                value_cell.border = thin_border
                if header in ["Total Count", "Bays Wide", "Bays Tall"]:
                    value_cell.alignment = Alignment(horizontal='left')
            
            output_section_current_row = 1
            profiles_for_section, accessories_for_section, other_items_for_section = [], [], []

            current_elevation_finish = elev_data.get("finish")

            for item in elev_data.get('calculated_outputs', []):
                pn, manual = item.get('part_number'), item.get('manual', False)
                if pn and pn != "N/A":
                    if manual:
                        other_items_for_section.append(item)
                    elif pn in PART_NUMBER_MAP.get("profiles", {}):
                        profiles_for_section.append(item)
                    elif pn in PART_NUMBER_MAP.get("accessories", {}) or item.get('type', '').lower() == 'accessory':
                        accessories_for_section.append(item)
                    else:
                        other_items_for_section.append(item)
                else:
                    other_items_for_section.append(item)

            system_total_for_this_block = [0.0]
            original_system_total_for_this_block = [0.0]
            newly_calculated_material_impacts_for_this_elevation = []

            next_row_after_profiles, impacts_p = _write_output_section(
                ws, "PROFILES", profiles_for_section, COL_E, current_elevation_finish,
                system_total_for_this_block, original_system_total_for_this_block, output_section_current_row,
                overall_current_extra_materials_state, private_extra_materials_path, multiplier
            )

            next_row_after_accessories, impacts_a = _write_output_section(
                ws, "ACCESSORIES", accessories_for_section, COL_E, current_elevation_finish,
                system_total_for_this_block, original_system_total_for_this_block, next_row_after_profiles,
                overall_current_extra_materials_state, private_extra_materials_path, multiplier
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
                    system_total_for_this_block, original_system_total_for_this_block, current_section_row,
                    overall_current_extra_materials_state, private_extra_materials_path, multiplier
                )
                newly_calculated_material_impacts_for_this_elevation.extend(impacts_g)
                current_section_row = next_row_after_group

            current_saved_elevations[elev_name]['material_impact'] = newly_calculated_material_impacts_for_this_elevation

            system_total_row = ws.max_row + 4
            
            lbl_col = PRICE_COL - 1
            val_col = PRICE_COL

            # Row 1: Original Total
            l_orig = ws.cell(row=system_total_row, column=lbl_col, value="Original Elevation Total")
            l_orig.font = Font(bold=True)
            l_orig.alignment = Alignment(horizontal='right')
            l_orig.border = Border(left=Side(style='thin'), top=Side(style='thin'))
            
            v_orig = ws.cell(row=system_total_row, column=val_col, value=original_system_total_for_this_block[0])
            v_orig.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            v_orig.font = Font(bold=True)
            v_orig.border = Border(right=Side(style='thin'), top=Side(style='thin'))

            # Row 2: Discounted Total
            l_disc = ws.cell(row=system_total_row + 1, column=lbl_col, value="Discounted Elevation Total")
            l_disc.font = Font(bold=True)
            l_disc.alignment = Alignment(horizontal='right')
            l_disc.border = Border(left=Side(style='thin'), bottom=Side(style='thin'))
            
            v_disc = ws.cell(row=system_total_row + 1, column=val_col, value=system_total_for_this_block[0])
            v_disc.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            v_disc.font = Font(bold=True)
            v_disc.border = Border(right=Side(style='thin'), bottom=Side(style='thin'))
            
            print(f"Rebuilt System Total for '{elev_name}': ${system_total_for_this_block[0]:.2f}")

            _autofit_columns(ws, COL_A, PRICE_COL, 1, ws.max_row)
            _clean_trailing_blank_rows(ws, 1)

    save_extra_materials(overall_current_extra_materials_state, private_extra_materials_path)

    summary_ws = wb.create_sheet(title="Summary")
    create_summary_sheet(summary_ws, private_elevations_path, private_extra_materials_path)
    
    final_save_path = os.path.join(public_reports_dir, os.path.basename(excel_path)) if mode == "export_all" else private_excel_path
    
    try:
        wb.save(final_save_path)
        print(f"Excel report '{final_save_path}' fully rebuilt with separate tabs.")
    except Exception as save_err:
        print(f"❌ Error saving Excel report during full rebuild: {save_err}")
        if completion_callback: completion_callback(f"Error saving report: {save_err}")
        return

    if mode != "export_all":
        try:
            with open(private_elevations_path, 'w') as f:
                json.dump(current_saved_elevations, f, indent=4)
        except IOError as e:
            print(f"Error saving all elevations to {private_elevations_path} after rebuild: {e}")

    if completion_callback:
        completion_callback()