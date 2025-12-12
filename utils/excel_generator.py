import os
import json
import math
from openpyxl import Workbook
from openpyxl.styles import Font, numbers, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import Counter
import datetime

from utils.pricing import get_price_by_part, reverse_material_impact, load_extra_materials, save_extra_materials, apply_material_impact_to_extra_materials_in_memory, get_unit_price_by_part, parse_length_to_feet, BAY_WIDTH_PARTS, _is_bay_width_part
EPSILON = 1e-9  # Small value for floating point comparisons
from data.part_number import PART_NUMBER_MAP
from data.parts_data import parts_data
from utils.formulas import calculate_door_info

# --- Helper Functions ---

def _get_multiplier(running_grand_total):
    """Returns multiplier based on running grand total."""
    return 0.614 if running_grand_total < 50000 else 0.572

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

def _create_bay_diagram(bays_wide, bays_tall, opening_width, opening_height, custom_bay_widths=None, custom_bay_heights=None):
    """
    Creates a visual blueprint diagram of the bay distribution.
    Returns a PIL Image object that can be inserted into Excel.
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
        import io
    except ImportError:
        print("PIL/Pillow not available, skipping diagram generation")
        return None
    
    # Diagram dimensions - smaller to fit in columns A-C without overlapping
    diagram_width = 400
    diagram_height = 300
    margin = 30
    
    # Calculate bay dimensions
    if custom_bay_widths and len(custom_bay_widths) == bays_wide:
        bay_widths = custom_bay_widths
    else:
        bay_widths = [opening_width / bays_wide] * bays_wide if bays_wide > 0 else []
    
    if custom_bay_heights and len(custom_bay_heights) == bays_tall:
        bay_heights = custom_bay_heights
    else:
        bay_heights = [opening_height / bays_tall] * bays_tall if bays_tall > 0 else []
    
    if not bay_widths or not bay_heights:
        return None
    
    # Create image
    img = Image.new('RGB', (diagram_width, diagram_height), color='white')
    draw = ImageDraw.Draw(img)
    
    # Calculate scaling to fit in diagram
    max_display_width = diagram_width - 2 * margin
    max_display_height = diagram_height - 2 * margin - 60  # Space for labels
    
    # Scale to fit
    total_width = sum(bay_widths)
    total_height = sum(bay_heights)
    scale_x = max_display_width / total_width if total_width > 0 else 1
    scale_y = max_display_height / total_height if total_height > 0 else 1
    scale = min(scale_x, scale_y)  # Maintain aspect ratio
    
    # Calculate starting position (centered)
    scaled_total_width = total_width * scale
    scaled_total_height = total_height * scale
    start_x = margin + (max_display_width - scaled_total_width) / 2
    start_y = margin + 30  # Space for title
    
    # Draw title - smaller fonts to prevent overlap
    try:
        font_large = ImageFont.truetype("arial.ttf", 12)
        font_small = ImageFont.truetype("arial.ttf", 8)
    except:
        try:
            font_large = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 12)
            font_small = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 8)
        except:
            font_large = ImageFont.load_default()
            font_small = ImageFont.load_default()
    
    draw.text((diagram_width // 2, 10), "Bay Distribution Layout", fill='black', anchor='mm', font=font_large)
    
    # Draw grid lines and bays
    current_x = start_x
    current_y = start_y
    
    # Draw vertical lines (bay width separators)
    for i, width in enumerate(bay_widths):
        if i > 0:
            draw.line([(current_x, start_y), (current_x, start_y + scaled_total_height)], fill='gray', width=2)
        current_x += width * scale
    
    # Draw horizontal lines (bay height separators)
    current_x = start_x
    for i, height in enumerate(bay_heights):
        if i > 0:
            draw.line([(start_x, current_y), (start_x + scaled_total_width, current_y)], fill='gray', width=2)
        current_y += height * scale
    
    # Draw outer border
    draw.rectangle([start_x, start_y, start_x + scaled_total_width, start_y + scaled_total_height], 
                   outline='black', width=3)
    
    # Draw bay labels
    current_x = start_x
    current_y = start_y
    bay_num = 1
    
    for row in range(bays_tall):
        current_x = start_x
        for col in range(bays_wide):
            # Calculate center of bay
            bay_center_x = current_x + (bay_widths[col] * scale) / 2
            bay_center_y = current_y + (bay_heights[row] * scale) / 2
            
            # Draw bay number - smaller spacing to prevent overlap
            draw.text((bay_center_x, bay_center_y - 6), f"B{bay_num}", fill='black', anchor='mm', font=font_small)
            
            # Draw dimensions - smaller spacing
            dim_text = f"{bay_widths[col]:.1f}\" x {bay_heights[row]:.1f}\""
            draw.text((bay_center_x, bay_center_y + 6), dim_text, fill='black', anchor='mm', font=font_small)
            
            current_x += bay_widths[col] * scale
            bay_num += 1
        current_y += bay_heights[row] * scale
    
    # Draw overall dimensions
    dim_text = f"Total: {opening_width:.1f}\" W x {opening_height:.1f}\" H"
    draw.text((diagram_width // 2, diagram_height - 20), dim_text, fill='black', anchor='mm', font=font_small)
    
    return img

def _add_bay_diagram_to_excel(ws, start_row, bays_wide, bays_tall, opening_width, opening_height, custom_bay_widths=None, custom_bay_heights=None):
    """Adds a bay distribution diagram to the Excel worksheet."""
    if bays_wide == 0 or bays_tall == 0:
        return start_row
    
    try:
        from openpyxl.drawing.image import Image as OpenpyxlImage
        import io
        
        # Create the diagram
        diagram_img = _create_bay_diagram(bays_wide, bays_tall, opening_width, opening_height, custom_bay_widths, custom_bay_heights)
        
        if diagram_img:
            # Save to bytes
            img_bytes = io.BytesIO()
            diagram_img.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            # Add to Excel - position in column A, spanning A-C but not overlapping column D
            img = OpenpyxlImage(img_bytes)
            # Resize image: fit within columns A-C (max width ~360 pixels to stay within 3 columns)
            # Original width was 180 pixels, so we'll use a width that fits in A-C
            original_width = img.width
            original_height = img.height
            # Maximum width to fit in columns A-C (approximately 360 pixels for 3 columns at width 20 each)
            max_width = 360
            img.width = min(450, max_width)  # Use larger size but cap at max_width to stay in A-C
            img.height = int(original_height * (img.width / original_width))  # Maintain aspect ratio
            img.anchor = f'A{start_row}'  # Place starting in column A
            ws.add_image(img)
            
            # Return the row after the image (estimate image height)
            # Image height in rows (approximately 1 row per 15 pixels at default row height)
            estimated_rows = max(15, int(img.height / 15))
            return start_row + estimated_rows + 2
    except Exception as e:
        print(f"Error creating bay diagram: {e}")
        # If diagram creation fails, just add a text note
        ws.cell(row=start_row, column=3, value="Bay diagram could not be generated")
    
    return start_row + 2

def _write_output_section(ws, title, items, colE, elevation_finish, system_total_ref, original_system_total_ref, start_output_row, current_extra_materials_state, extra_materials_path, multiplier, show_qty_per_elevation=False, total_count=1, show_total_cost_per_elevation=False, show_discounted_cost_per_elevation=False):
    """Writes a section of calculated outputs to the worksheet."""
    if not items: 
        return start_output_row, [], {'original': 0.0, 'discounted': 0.0}

    current_row = start_output_row
    title_cell = ws.cell(row=current_row, column=colE, value=title)
    title_cell.font = Font(bold=True, size=12)
    # title_cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # Removed color fill for professional look

    # Build headers based on which optional columns to show
    headers = ["Description", "Part Number", "Total Quantity Required"]
    if show_qty_per_elevation and total_count > 1:
        headers.append("Quantity Per Elevation")
    headers.append("Total List Cost")
    if show_total_cost_per_elevation and total_count > 1:
        headers.append("Total List Cost Per Elevation")
    headers.append("Discounted Total List Cost")
    if show_discounted_cost_per_elevation and total_count > 1:
        headers.append("Discounted Total List Cost Per Elevation")
    
    for i, h in enumerate(headers):
        header_cell = ws.cell(row=current_row + 1, column=colE + i, value=h)
        header_cell.font = Font(bold=True)
        header_cell.border = Border(bottom=Side(style='thin'))
        # header_cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid") # Removed color fill for professional look
    current_row += 2

    section_material_impacts = []
    section_original_total = 0.0
    section_discounted_total = 0.0
    section_original_per_elev_total = 0.0  # Sum of actual per-elevation costs
    section_discounted_per_elev_total = 0.0  # Sum of actual per-elevation discounted costs

    for item in items:
        qty_raw = item.get('quantity', 0)
        pn, manual = item.get('part_number'), item.get('manual', False)
        desc = item.get('description', '').strip()
        is_profile = pn in PART_NUMBER_MAP.get('profiles', {})
        is_gasket = "gasket" in desc.lower() or pn in ["E2-0052", "E2-0053", "E2-0065"]
        is_accessory = pn in PART_NUMBER_MAP.get('accessories', {}) or item.get('type', '').lower() == 'accessory'
        is_glass = pn == "GLASS_AREA" or item.get('type', '').lower() == 'glass'

        # Determine if we should process as a group (for optimization)
        is_bay_width_item = _is_bay_width_part(pn, qty_raw, item.get('description', ''))
        is_list = isinstance(qty_raw, list)
        has_multiple_items = is_list and len(qty_raw) > 1
        
        # For profiles/gaskets with a list, ALWAYS process the entire list at once for optimization
        # This ensures proper leftover calculation across multiple cuts
        # CRITICAL: Check this BEFORE creating individual_quantities to avoid splitting the list
        should_process_as_group = (is_profile or is_gasket) and has_multiple_items
        
        # Calculate quantities for display and processing
        # Only split into individual_quantities if NOT processing as group
        if should_process_as_group:
            # Keep as list for group processing
            individual_quantities = qty_raw if is_list else [qty_raw]
            qty_sum = sum(qty_raw) if is_list else qty_raw
        else:
            # Split for individual processing
            individual_quantities = qty_raw if is_list else [qty_raw]
            qty_sum = sum(individual_quantities) if is_list else qty_raw

        unit_type = 'ft' if (is_profile or is_gasket) else 'pcs' if is_accessory else item.get('unit', 'pcs' if not is_glass else 'sqft')
        display_unit = unit_type

        # Format display string
        if is_list:
            if len(qty_raw) > 1 and all(x == qty_raw[0] for x in qty_raw):
                # For profiles, show as "8ft x 3" format, for others use decimal format
                if is_profile:
                    # Check if it's a whole number
                    if qty_raw[0] == int(qty_raw[0]):
                        display_qty_string = f"{int(qty_raw[0])}{display_unit} x {len(qty_raw)}"
                    else:
                        display_qty_string = f"{qty_raw[0]:.2f}{display_unit} x {len(qty_raw)}"
                else:
                    display_qty_string = f"{qty_raw[0]:.2f} {display_unit} x {len(qty_raw)}"
            else:
                # For profiles, show individual cuts without decimals when whole numbers, with decimals otherwise
                if is_profile:
                    display_qty_string = ", ".join([f"{int(q)}{display_unit}" if q == int(q) else f"{q:.2f}{display_unit}" for q in qty_raw])
                else:
                    display_qty_string = ", ".join([f"{q:.2f} {display_unit}" for q in qty_raw])
        else:
            # For profiles, show without decimals when whole number
            if is_profile and qty_raw == int(qty_raw):
                display_qty_string = f"{int(qty_raw)}{display_unit}"
            else:
                display_qty_string = f"{qty_raw:.2f} {display_unit}"

        item_total_cost_for_display = 0.0
        original_item_total_cost = 0.0
        
        # Process as group if it's a profile/gasket with multiple pieces
        if should_process_as_group:
            # Process the entire list as one request for waste optimization
            # Profiles and gaskets are treated as length-based items (sold by length, with leftover tracking)
            print(f"DEBUG: Processing {pn} as GROUP with list {qty_raw} (type: {type(qty_raw)}, is_list: {isinstance(qty_raw, list)})")
            use_group = is_profile or is_gasket
            # CRITICAL: Ensure qty_raw is passed as a list, not converted to a single value
            list_to_process = qty_raw if isinstance(qty_raw, list) else [qty_raw]
            print(f"DEBUG: Calling get_price_by_part with list: {list_to_process}")
            total_price, calculated_unit_type, material_impact_details = \
                get_price_by_part(pn, list_to_process, finish=elevation_finish, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False, group=use_group, description=item.get('description', ''))
            
            item_total_cost_for_display = total_price or 0.0
            original_item_total_cost = total_price or 0.0
            calculated_unit_type = unit_type if (is_profile or is_gasket or is_accessory) else (calculated_unit_type or item.get('unit', 'pcs'))
            
            if material_impact_details:
                leftover_qty = material_impact_details.get('leftover_generated_qty_or_length', 0.0)
                all_leftovers = material_impact_details.get('all_new_leftovers', [])
                print(f"DEBUG: Material impact for {pn}: leftover_qty={leftover_qty}, all_new_leftovers={all_leftovers}")
                material_impact_details['leftover_generated_qty_or_length_display'] = f"{leftover_qty:.2f} {display_unit}"
                section_material_impacts.append(material_impact_details)
                apply_material_impact_to_extra_materials_in_memory(current_extra_materials_state, material_impact_details)
            else:
                print(f"DEBUG: WARNING - No material_impact_details returned for {pn} with list {qty_raw}")
        else:
            # Standard processing: iterate through each quantity (for non-profile/gasket items or single quantities)
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
                    # For profiles/gaskets, use group=True to ensure proper processing
                    # For others, use group only if it's a gasket
                    use_group = (is_profile or is_gasket)
                    total_price, unit_from_pricing, material_impact_details = \
                        get_price_by_part(pn, single_qty_for_calc, finish=elevation_finish, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False, group=use_group)
                    total_item_price_single_cut = total_price or 0.0
                    calculated_unit_type = unit_type if (is_profile or is_gasket or is_accessory) else (unit_from_pricing or item.get('unit', 'pcs'))

                item_total_cost_for_display += total_item_price_single_cut
                original_item_total_cost += total_item_price_single_cut

                if material_impact_details:
                    leftover_qty = material_impact_details.get('leftover_generated_qty_or_length', 0.0)
                    material_impact_details['leftover_generated_qty_or_length_display'] = f"{leftover_qty:.2f} {display_unit}"
                    section_material_impacts.append(material_impact_details)
                    apply_material_impact_to_extra_materials_in_memory(current_extra_materials_state, material_impact_details)

        if is_profile or is_gasket or is_accessory:
            item_total_cost_for_display *= multiplier
            if qty_sum > 0:
                item['price'] = item_total_cost_for_display / qty_sum

        system_total_ref[0] += item_total_cost_for_display
        original_system_total_ref[0] += original_item_total_cost
        section_original_total += original_item_total_cost
        section_discounted_total += item_total_cost_for_display

        # Calculate quantity per elevation if needed - show as "8ft x 2" format (pieces per elevation) for profiles/gaskets only
        qty_per_elev_display = None
        if show_qty_per_elevation and total_count > 1:
            # Only use "x" format for profiles and gaskets (they have specific cut dimensions)
            use_x_format = is_profile or is_gasket
            
            if isinstance(qty_raw, list):
                num_pieces = len(qty_raw)
                pieces_per_elev = num_pieces / total_count
                
                # If all pieces are the same length, show as "8ft x 2" for profiles/gaskets
                if len(qty_raw) > 1 and all(abs(x - qty_raw[0]) < 0.001 for x in qty_raw):
                    piece_length = qty_raw[0]
                    if use_x_format:
                        if is_profile:
                            if piece_length == int(piece_length):
                                qty_per_elev_display = f"{int(piece_length)}{display_unit} x {int(pieces_per_elev)}"
                            else:
                                qty_per_elev_display = f"{piece_length:.2f}{display_unit} x {int(pieces_per_elev)}"
                        else:  # is_gasket
                            qty_per_elev_display = f"{piece_length:.2f} {display_unit} x {int(pieces_per_elev)}"
                    else:
                        # For other items, just show the quantity per elevation without "x"
                        qty_per_elev_display = f"{pieces_per_elev:.2f} {display_unit}"
                else:
                    # Different lengths - group by length and show pieces per elevation for each
                    if use_x_format:
                        length_counts = Counter(qty_raw)
                        parts = []
                        for length, count in sorted(length_counts.items()):
                            pieces_per_length_per_elev = (count / total_count)
                            if is_profile:
                                if length == int(length):
                                    parts.append(f"{int(length)}{display_unit} x {int(pieces_per_length_per_elev)}")
                                else:
                                    parts.append(f"{length:.2f}{display_unit} x {int(pieces_per_length_per_elev)}")
                            else:  # is_gasket
                                parts.append(f"{length:.2f} {display_unit} x {int(pieces_per_length_per_elev)}")
                        qty_per_elev_display = ", ".join(parts)
                    else:
                        # For other items, just show total quantity per elevation
                        qty_per_elev = qty_sum / total_count
                        qty_per_elev_display = f"{qty_per_elev:.2f} {display_unit}"
            else:
                # Single quantity value
                qty_per_elev = qty_raw / total_count
                if use_x_format:
                    # For profiles/gaskets, show as "8ft x 1"
                    if is_profile:
                        if qty_raw == int(qty_raw):
                            qty_per_elev_display = f"{int(qty_raw)}{display_unit} x 1"
                        else:
                            qty_per_elev_display = f"{qty_raw:.2f}{display_unit} x 1"
                    else:  # is_gasket
                        qty_per_elev_display = f"{qty_raw:.2f} {display_unit} x 1"
                else:
                    # For other items, just show the quantity per elevation without "x"
                    qty_per_elev_display = f"{qty_per_elev:.2f} {display_unit}"
        
        # Write data columns
        col_offset = 0
        ws.cell(row=current_row, column=colE + col_offset, value=item.get('description', ''))
        col_offset += 1
        ws.cell(row=current_row, column=colE + col_offset, value=pn or 'N/A')
        col_offset += 1
        ws.cell(row=current_row, column=colE + col_offset, value=display_qty_string)
        col_offset += 1
        
        # Add quantity per elevation column if enabled
        if qty_per_elev_display:
            ws.cell(row=current_row, column=colE + col_offset, value=qty_per_elev_display)
            col_offset += 1
        
        # Total List Cost
        ws.cell(row=current_row, column=colE + col_offset, value=original_item_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        col_offset += 1
        
        # Total List Cost Per Elevation (if enabled) - calculate actual purchase cost for per-elevation quantity
        original_cost_per_elev = None
        if show_total_cost_per_elevation and total_count > 1:
            # Calculate the quantity per elevation
            if isinstance(qty_raw, list):
                # For lists, calculate pieces per elevation
                num_pieces = len(qty_raw)
                pieces_per_elev = num_pieces / total_count
                
                # Check if this is a bay width part - if so, keep as list
                is_bay_width_for_per_elev = _is_bay_width_part(pn, qty_raw, item.get('description', ''))
                
                if is_bay_width_for_per_elev:
                    # For bay width parts, take first pieces_per_elev items
                    pieces_per_elev_int = int(pieces_per_elev)
                    if pieces_per_elev_int > 0 and pieces_per_elev_int <= len(qty_raw):
                        qty_per_elev = qty_raw[:pieces_per_elev_int]
                    else:
                        # Fallback: divide each piece equally (not ideal but works)
                        qty_per_elev = [q / total_count for q in qty_raw]
                else:
                    # For non-bay-width parts, if all pieces are the same, use single value
                    if len(qty_raw) > 0 and all(abs(x - qty_raw[0]) < 0.001 for x in qty_raw):
                        # All pieces same length, per elevation is just one piece
                        qty_per_elev = qty_raw[0]
                    else:
                        # Different lengths - sum and divide
                        qty_per_elev = qty_sum / total_count
            else:
                # Single quantity value
                qty_per_elev = qty_raw / total_count
            
            # Calculate actual purchase cost for per-elevation quantity (accounts for minimum purchase lengths)
            if not manual and pn and pn != "N/A":
                use_group_for_per_elev = is_gasket
                
                # Get price for per-elevation quantity
                per_elev_price, _, _ = get_price_by_part(
                    pn, qty_per_elev, 
                    finish=elevation_finish, 
                    current_extra_materials=current_extra_materials_state, 
                    extra_materials_file=extra_materials_path, 
                    summary=True,  # Use summary=True to get price without material impact
                    group=use_group_for_per_elev,
                    description=item.get('description', '')
                )
                original_cost_per_elev = per_elev_price if per_elev_price is not None else (original_item_total_cost / total_count)
            else:
                # For manual items, just divide the cost
                original_cost_per_elev = original_item_total_cost / total_count
            
            ws.cell(row=current_row, column=colE + col_offset, value=original_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            section_original_per_elev_total += original_cost_per_elev  # Track sum of per-elevation costs
            col_offset += 1
        
        # Discounted Total List Cost
        ws.cell(row=current_row, column=colE + col_offset, value=item_total_cost_for_display).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        col_offset += 1
        
        # Discounted Total List Cost Per Elevation (if enabled) - apply multiplier to per-elevation cost
        if show_discounted_cost_per_elevation and total_count > 1:
            if original_cost_per_elev is not None:
                # Use the calculated per-elevation cost and apply multiplier
                discounted_cost_per_elev = original_cost_per_elev * multiplier if (is_profile or is_gasket or is_accessory) else original_cost_per_elev
            else:
                # Fallback: recalculate per-elevation quantity and price (shouldn't happen if original_cost_per_elev was calculated)
                if isinstance(qty_raw, list):
                    num_pieces = len(qty_raw)
                    pieces_per_elev = num_pieces / total_count
                    is_bay_width_for_per_elev = _is_bay_width_part(pn, qty_raw, item.get('description', ''))
                    
                    if is_bay_width_for_per_elev:
                        pieces_per_elev_int = int(pieces_per_elev)
                        if pieces_per_elev_int > 0 and pieces_per_elev_int <= len(qty_raw):
                            qty_per_elev = qty_raw[:pieces_per_elev_int]
                        else:
                            qty_per_elev = [q / total_count for q in qty_raw]
                    else:
                        if len(qty_raw) > 0 and all(abs(x - qty_raw[0]) < 0.001 for x in qty_raw):
                            qty_per_elev = qty_raw[0]
                        else:
                            qty_per_elev = qty_sum / total_count
                else:
                    qty_per_elev = qty_raw / total_count
                
                if not manual and pn and pn != "N/A":
                    use_group_for_per_elev = is_gasket
                    per_elev_price, _, _ = get_price_by_part(
                        pn, qty_per_elev,
                        finish=elevation_finish,
                        current_extra_materials=current_extra_materials_state,
                        extra_materials_file=extra_materials_path,
                        summary=True,
                        group=use_group_for_per_elev,
                        description=item.get('description', '')
                    )
                    discounted_cost_per_elev = (per_elev_price * multiplier) if (is_profile or is_gasket or is_accessory) and per_elev_price is not None else (item_total_cost_for_display / total_count)
                else:
                    discounted_cost_per_elev = item_total_cost_for_display / total_count
            ws.cell(row=current_row, column=colE + col_offset, value=discounted_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            section_discounted_per_elev_total += discounted_cost_per_elev  # Track sum of per-elevation discounted costs
            col_offset += 1
        
        current_row += 1

    # Add Section Totals
    # Convert title to proper label (e.g., "PROFILES" -> "Profile Cost")
    title_mapping = {
        "PROFILES": "Profile",
        "ACCESSORIES": "Accessory",
        "GASKETS": "Gasket",
        "DOORS": "Door",
        "GLASS": "Glass",
        "LABOR": "Labor"
    }
    title_label = title_mapping.get(title.upper(), title.title())
    total_label = f"Total {title_label} Cost"
    
    # Calculate column offsets based on which optional columns are shown
    total_col_offset = 2  # Start after "Total Quantity Required"
    if show_qty_per_elevation and total_count > 1:
        total_col_offset += 1  # Skip "Quantity Per Elevation" column
    
    # Place total label in the "Total List Cost" column
    ws.cell(row=current_row, column=colE + total_col_offset, value=total_label).font = Font(bold=True)
    total_col_offset += 1
    
    # Total List Cost
    ws.cell(row=current_row, column=colE + total_col_offset, value=section_original_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=current_row, column=colE + total_col_offset).font = Font(bold=True)
    total_col_offset += 1
    
    # Total List Cost Per Elevation (if enabled) - use sum of actual per-elevation costs
    if show_total_cost_per_elevation and total_count > 1:
        ws.cell(row=current_row, column=colE + total_col_offset, value=section_original_per_elev_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=colE + total_col_offset).font = Font(bold=True)
        total_col_offset += 1
    
    # Discounted Total List Cost
    ws.cell(row=current_row, column=colE + total_col_offset, value=section_discounted_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=current_row, column=colE + total_col_offset).font = Font(bold=True)
    total_col_offset += 1
    
    # Discounted Total List Cost Per Elevation (if enabled) - use sum of actual per-elevation discounted costs
    if show_discounted_cost_per_elevation and total_count > 1:
        ws.cell(row=current_row, column=colE + total_col_offset, value=section_discounted_per_elev_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=colE + total_col_offset).font = Font(bold=True)
        total_col_offset += 1
    
    # Add top border for totals row - adjust range based on number of columns
    num_cols = 5  # Base columns: Description, Part Number, Total Quantity Required, Total List Cost, Discounted Total List Cost
    if show_qty_per_elevation and total_count > 1:
        num_cols += 1
    if show_total_cost_per_elevation and total_count > 1:
        num_cols += 1
    if show_discounted_cost_per_elevation and total_count > 1:
        num_cols += 1
    for col in range(colE, colE + num_cols):
        ws.cell(row=current_row, column=col).border = Border(top=Side(style='thin'))

    # Return section totals for summary
    section_totals = {
        'original': section_original_total,
        'discounted': section_discounted_total
    }
    return current_row + 1, section_material_impacts, section_totals


def create_summary_sheet(ws, elevations_json_path, extra_materials_json_path, wb=None, summary_settings_path=None):
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

    # For summary, use a shared in-memory state that accumulates leftovers across all elevations
    # This allows the summary to utilize waste materials across all elevations
    summary_extra_materials_state = {}
    
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
        'GASKETS': [],
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
            is_gasket = "gasket" in desc.lower() or part in ["E2-0052", "E2-0053", "E2-0065"]  # Identify gaskets
            is_accessory = part in PART_NUMBER_MAP.get('accessories', {}) or output.get('type', '').lower() == 'accessory'
            is_glass = part == "GLASS_AREA" or output.get('type', '').lower() == 'glass'
            is_joints_fab_labor = part == "JOINTS_FAB_LABOR" or output.get('type', '').lower() == 'joints_fab_labor' or "joints fabrication" in desc.lower() or "fabrication labor" in desc.lower()
            is_door = output.get('type', '').lower() in ['door', 'doors']

            if is_profile:
                category = 'PROFILES'
            elif is_gasket:
                category = 'GASKETS'
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
                    key = f"MANUAL_{part}-{elevation_finish}" if (is_profile or is_gasket or is_joints_fab_labor or is_door or is_glass) and elevation_finish else f"MANUAL_{part}"
                    display = f"{desc} ({part} - {elevation_finish})" if (is_profile or is_gasket or is_joints_fab_labor or is_door or is_glass) and elevation_finish else f"{desc} ({part})"
                else:
                    key = f"MANUAL_NO_PN_{desc}"
                    display = desc
            else:
                if (is_profile or is_gasket or is_joints_fab_labor or is_door or is_glass) and elevation_finish:
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
                'unit': 'ft' if (is_profile or is_gasket) else 'pcs' if is_accessory else output.get('unit', 'pcs' if not is_glass else 'sqft'),
                'finish': elevation_finish if (is_profile or is_gasket or is_joints_fab_labor or is_door or is_glass) else '',
                'is_glass': is_glass,
                'is_joints_fab_labor': is_joints_fab_labor,
                'is_door': is_door,
                'is_gasket': is_gasket,
                'price': output.get('price', 0.0) if (manual or is_glass or is_joints_fab_labor or is_door) else 0.0
            })

    # Step 2.5: Aggregate items within each category by key to prevent duplicates across elevations
    for category in categories:
        aggregated_map = {}
        for item in categories[category]:
            k = item['key']
            if k in aggregated_map:
                existing = aggregated_map[k]
                # Preserve description if missing
                if not existing.get('description') and item.get('description'):
                    existing['description'] = item['description']
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
                # Ensure description exists
                if 'description' not in item or not item.get('description'):
                    item['description'] = item.get('display', '')
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
            description = item.get('description', '') or display
            is_profile = part in PART_NUMBER_MAP.get('profiles', {})
            is_accessory = part in PART_NUMBER_MAP.get('accessories', {}) or item.get('type', '').lower() == 'accessory'
            is_gasket = item.get('is_gasket', False) or "gasket" in item.get('description', '').lower() or part in ["E2-0052", "E2-0053", "E2-0065"]
            is_glass = item['is_glass']
            is_joints_fab_labor = item['is_joints_fab_labor']
            is_door = item['is_door']
            item_finish = item['finish']

            display_unit = 'ft' if (is_profile or is_gasket) else 'pcs' if is_accessory else item['unit']
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
                # Gaskets are treated as profiles (sold by length, with leftover tracking)
                use_group = is_gasket
                
                # For profiles and gaskets, use quantity_list if available for optimal cutting
                # This allows the pricing function to optimize cuts across multiple pieces
                quantity_for_pricing = quantity_aggregated
                if (is_profile or is_gasket) and part and part != "N/A":
                    quantity_list = item.get('quantity_list', [])
                    if quantity_list and len(quantity_list) > 1:
                        # Use the list for optimization - filter valid values
                        valid_quantities = [q for q in quantity_list if q is not None and isinstance(q, (int, float)) and q > 0]
                        if valid_quantities:
                            quantity_for_pricing = valid_quantities
                
                # Use shared in-memory state for summary to accumulate leftovers across all elevations
                total_price, unit_type_from_pricing, material_impact = get_price_by_part(
                    part,
                    quantity_for_pricing,
                    finish=item_finish,
                    current_extra_materials=summary_extra_materials_state,
                    extra_materials_file=extra_materials_json_path,
                    summary=False,  # Use False to actually use and track leftovers
                    group=use_group
                )
                # Apply material impact to accumulate leftovers in summary state
                if material_impact:
                    apply_material_impact_to_extra_materials_in_memory(summary_extra_materials_state, material_impact)
                original_total_cost_for_item = total_price if total_price is not None else 0.0
                calculated_unit_type = 'ft' if is_profile else 'pcs' if is_accessory else (unit_type_from_pricing or item['unit'] or 'pcs')

            if is_profile or is_gasket or is_accessory:
                total_cost_for_item = original_total_cost_for_item * multiplier
            else:
                total_cost_for_item = original_total_cost_for_item

            total_discounted_price += total_cost_for_item

            if part and part != "N/A" and (is_profile or is_gasket or is_accessory):
                extra_materials_key_for_reuse = part
                if (is_profile or is_gasket) and item_finish:
                    extra_materials_key_for_reuse = f"{part}-{item_finish}"

                # Use the shared summary state that accumulates leftovers across all elevations
                part_data = summary_extra_materials_state.get(extra_materials_key_for_reuse, {})
                if part_data.get("length_pieces"):
                    # Get min_purchase_length for validation
                    part_info = parts_data.get(part, {})
                    length_str = part_info.get('Length', '')
                    min_purchase_length = parse_length_to_feet(length_str) or 24.0
                    
                    # Filter out invalid leftover pieces (must be > 0 and < min_purchase_length)
                    lengths = [float(x) for x in part_data["length_pieces"] if isinstance(x, (int, float, str))]
                    valid_lengths = [l for l in lengths if l > EPSILON and l < min_purchase_length - EPSILON]
                    reusable_qty_sum = sum(valid_lengths)
                    if valid_lengths:
                        # Count occurrences of each length, using rounded values for grouping
                        counter = Counter([round(l, 2) for l in valid_lengths])
                        reuse_lengths_formatted = []
                        for length_val, count in sorted(counter.items(), key=lambda x: x[0], reverse=True):
                            # Format length: use integer if whole number, otherwise 2 decimals
                            if length_val == int(length_val):
                                length_str = f"{int(length_val)}{display_unit}"
                            else:
                                length_str = f"{length_val:.2f}{display_unit}"
                            # Always show count, even if it's 1
                            reuse_lengths_formatted.append(f"{length_str} x{count}")
                        reusable_qty_display_string = ", ".join(reuse_lengths_formatted)
                    else:
                        reusable_qty_sum = 0.0
                        reusable_qty_display_string = "N/A"
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
            # Calculate Quantity Req (FT) and Qty Stick/Roll (Req)
            quantity_req_ft = "N/A"
            qty_stick_req = "N/A"
            quantity_display_formatted = f"{quantity_aggregated:.2f} {display_unit}"
            
            if (is_profile or is_gasket) and part and part != "N/A":
                # For profiles and gaskets, quantity is in feet - add units
                quantity_req_ft = f"{quantity_aggregated:.2f} ft"
                # Calculate number of sticks/rolls needed and format with length
                part_data = parts_data.get(part, {})
                length_str = part_data.get('Length', '')
                min_purchase_length = parse_length_to_feet(length_str) or 1.0
                if min_purchase_length > 0:
                    num_units = math.ceil(quantity_aggregated / min_purchase_length)
                    unit_label = "rolls" if is_gasket else "sticks"
                    qty_stick_req = f"{num_units} ({min_purchase_length:.0f}ft per)"
                else:
                    qty_stick_req = "N/A"
                
                # Format quantity_display to show breakdown like "16ft x2, 8ft x1"
                quantity_list = item.get('quantity_list', [quantity_aggregated])
                # Filter out invalid values (None, empty, non-numeric)
                valid_quantities = [q for q in quantity_list if q is not None and (isinstance(q, (int, float)) and q > 0)]
                
                # If no valid quantities, use aggregated quantity
                if not valid_quantities:
                    valid_quantities = [quantity_aggregated] if quantity_aggregated > 0 else []
                
                # Count occurrences of each length
                if valid_quantities:
                    length_counter = Counter([round(q, 2) for q in valid_quantities])
                    if len(length_counter) > 1:
                        # Multiple different lengths - show breakdown
                        length_parts = []
                        for length_val, count in sorted(length_counter.items(), key=lambda x: x[0], reverse=True):
                            if count > 1:
                                length_parts.append(f"{length_val:.0f}ft x{count}")
                            else:
                                length_parts.append(f"{length_val:.0f}ft x1")
                        quantity_display_formatted = ", ".join(length_parts)
                    elif len(length_counter) == 1:
                        # All same length - show total with count if multiple pieces
                        length_val = list(length_counter.keys())[0]
                        count = length_counter[length_val]
                        if count > 1:
                            quantity_display_formatted = f"{length_val:.0f}ft x{count}"
                        else:
                            quantity_display_formatted = f"{length_val:.0f}ft"
                    else:
                        # Empty counter - fallback to default format
                        quantity_display_formatted = f"{quantity_aggregated:.2f} {display_unit}"
                else:
                    # No valid quantities - fallback to default format
                    quantity_display_formatted = f"{quantity_aggregated:.2f} {display_unit}"
                    
            elif is_accessory and part and part != "N/A":
                # For accessories, get bulk order info
                part_data = parts_data.get(part, {})
                units_str = part_data.get('Units', '1 pcs.')
                length_str = part_data.get('Length', '')
                
                # Check if sold by length (feet) or by pieces
                length_ft = parse_length_to_feet(length_str) if length_str else 0.0
                
                if length_ft > 1.0:
                    # Sold by length (e.g., glazing gasket: 500 ft per order)
                    unit_count_per_bundle = length_ft
                    unit_label = "ft per"
                else:
                    # Sold by pieces
                    unit_count_per_bundle = 1
                    if 'pc' in units_str.lower():
                        try:
                            unit_count_per_bundle = int(units_str.lower().split('pc')[0].strip()) or 1
                        except ValueError:
                            unit_count_per_bundle = 1
                    unit_label = "pcs per"
                
                # Column 2: Quantity (total pieces/feet required)
                quantity_req_ft = f"{quantity_aggregated:.2f} {display_unit}"
                
                # Column 3: Quantity Per Order (e.g., "500 ft per" or "20 pcs per")
                qty_stick_req = f"{unit_count_per_bundle:.0f} {unit_label}"
                
                # Column 4: Orders Required (number of orders needed)
                num_orders = math.ceil(quantity_aggregated / unit_count_per_bundle) if unit_count_per_bundle > 0 else 0
                quantity_display_formatted = f"{num_orders} order{'s' if num_orders != 1 else ''}"
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
                'description': description,
                'display': display,
                'quantity_display': quantity_display_formatted,
                'quantity_req_ft': quantity_req_ft,
                'qty_stick_req': qty_stick_req,
                'original_total_cost': original_total_cost_for_item,
                'total_cost': total_cost_for_item,
                'reusable_qty_display': reusable_qty_display_string,
                'reusable_pct': reusable_pct if (is_profile or is_gasket or is_accessory) else "N/A",
                'reusable_cost': reusable_cost if (is_profile or is_gasket or is_accessory) else 0.0,
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
                "Description", "Project Total Materials", "Total Feet Required", "Sticks Required", "Total Quantity Required", "Total List Cost", "Discounted Total List Cost",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'ACCESSORIES':
            return [
                "Description", "Project Total Materials", "Total Pieces Required", "Quantity Per Order", "Orders Required", "Total List Cost", "Discounted Total List Cost",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'GASKETS':
            return [
                "Description", "Project Total Materials", "Total Feet Required", "Rolls Required", "Total Quantity Required", "Total List Cost", "Discounted Total List Cost",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'GLASS':
            return [
                "Description", "Project Total Materials", "N/A", "Unit Price", "Total Quantity Required", "Total List Cost", "Discounted Total List Cost",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'LABOR':
            return [
                "Description", "Project Total Materials", "N/A", "Unit Price", "Total Quantity Required", "Total List Cost", "Discounted Total List Cost",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        elif category == 'DOORS':
            return [
                "Description", "Project Total Materials", "N/A", "Unit Price", "Total Quantity Required", "Total List Cost", "Discounted Total List Cost",
                "Residual Material Quantity", "Residual Waste %", "Residual Material Cost"
            ]
        else:
            # Default headers
            return [
                "Description", "Project Total Materials", "Quantity Req (FT)", "Qty Stick (Req)", "Total Quantity Required", "Total List Cost", "Discounted Total List Cost",
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
                
                # Get description - it should be in final_summary_data
                description_value = item.get('description', '')
                if not description_value or description_value == '':
                    # Fallback to display if description is missing
                    description_value = item.get('display', '')
                ws.cell(row=current_row, column=1, value=description_value)
                ws.cell(row=current_row, column=2, value=item['display'])
                ws.cell(row=current_row, column=3, value=item['quantity_req_ft'])
                ws.cell(row=current_row, column=4, value=item['qty_stick_req'])
                ws.cell(row=current_row, column=5, value=item['quantity_display'])
                ws.cell(row=current_row, column=6, value=item['original_total_cost']).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                ws.cell(row=current_row, column=7, value=item['total_cost']).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                ws.cell(row=current_row, column=8, value=item['reusable_qty_display'])
                ws.cell(row=current_row, column=9, value=f"{item['reusable_pct']:.2f}%" if isinstance(item['reusable_pct'], (int, float)) else item['reusable_pct'])
                ws.cell(row=current_row, column=10, value=item['reusable_cost']).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                current_row += 1
        
        grand_original_total += section_original_total
        grand_discounted_total += section_total_cost
        grand_residual_total += section_residual_total

        # Add Section Totals for Summary
        # Convert category to proper label (e.g., "PROFILES" -> "Profile Cost")
        category_mapping = {
            "PROFILES": "Profile",
            "ACCESSORIES": "Accessory",
            "GASKETS": "Gasket",
            "DOORS": "Door",
            "GLASS": "Glass",
            "LABOR": "Labor"
        }
        category_label = category_mapping.get(category.upper(), category.title())
        total_label = f"Total {category_label} Cost"
        ws.cell(row=current_row, column=5, value=total_label).font = Font(bold=True)
        ws.cell(row=current_row, column=6, value=section_original_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=6).font = Font(bold=True)
        ws.cell(row=current_row, column=7, value=section_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=7).font = Font(bold=True)
        ws.cell(row=current_row, column=10, value=section_residual_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=10).font = Font(bold=True)
        
        # Add top border for totals row
        for col in range(1, 11):
            ws.cell(row=current_row, column=col).border = Border(top=Side(style='thin'))

        current_row += 2

    # Grand Totals Block
    gt_row = current_row + 2
    
    # Original Total
    ws.cell(row=gt_row, column=6, value="Overall Total Price (List)").font = Font(bold=True)
    ws.cell(row=gt_row, column=6).alignment = Alignment(horizontal='right')
    ws.cell(row=gt_row, column=6).border = Border(left=Side(style='thin'), top=Side(style='thin'))
    
    ws.cell(row=gt_row, column=7, value=grand_original_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=gt_row, column=7).font = Font(bold=True)
    ws.cell(row=gt_row, column=7).border = Border(right=Side(style='thin'), top=Side(style='thin'))

    # Discounted Total - sum column G (column 7) from all section total rows in the summary
    # Column 7 is "Discounted Total List Cost" and contains the section totals
    sum_from_column_g = 0.0
    try:
        # Find all section total rows - they have "Total X Cost" in column 5 and values in column 7
        for row in range(1, gt_row):
            label_cell = ws.cell(row=row, column=5)  # Column E (5) contains labels like "Total Profile Cost"
            value_cell = ws.cell(row=row, column=7)  # Column G (7) contains the discounted total cost
            
            if label_cell.value and isinstance(label_cell.value, str):
                if "Total" in label_cell.value and "Cost" in label_cell.value:
                    # This is a section total row
                    if value_cell.value is not None:
                        try:
                            sum_from_column_g += float(value_cell.value)
                            print(f"Found section total '{label_cell.value}' in row {row}, column 7: ${value_cell.value}")
                        except (ValueError, TypeError):
                            pass
    except Exception as e:
        print(f"Error reading from column G: {e}")
        import traceback
        traceback.print_exc()
        sum_from_column_g = 0.0
    
    # Use sum from column G if available, otherwise use calculated total
    final_discounted_total = sum_from_column_g if sum_from_column_g > 0 else grand_discounted_total
    if sum_from_column_g > 0:
        print(f"Summary discounted total from column G: ${sum_from_column_g:.2f}, calculated: ${grand_discounted_total:.2f}, using: ${final_discounted_total:.2f}")
    
    ws.cell(row=gt_row+1, column=6, value="Overall Discounted Total").font = Font(bold=True)
    ws.cell(row=gt_row+1, column=6).alignment = Alignment(horizontal='right')
    ws.cell(row=gt_row+1, column=6).border = Border(left=Side(style='thin'))

    ws.cell(row=gt_row+1, column=7, value=final_discounted_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=gt_row+1, column=7).font = Font(bold=True)
    ws.cell(row=gt_row+1, column=7).border = Border(right=Side(style='thin'))

    # Residual Cost
    reuse_total = total_reusable_cost
    ws.cell(row=gt_row+2, column=6, value="Overall Residual Cost").font = Font(bold=True)
    ws.cell(row=gt_row+2, column=6).alignment = Alignment(horizontal='right')
    ws.cell(row=gt_row+2, column=6).border = Border(left=Side(style='thin'))

    ws.cell(row=gt_row+2, column=7, value=reuse_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    ws.cell(row=gt_row+2, column=7).font = Font(bold=True)
    ws.cell(row=gt_row+2, column=7).border = Border(right=Side(style='thin'))

    # Waste %
    reuse_pct_of_gt = min((total_reusable_cost / total_discounted_price * 100) if total_discounted_price > 0 else 0.0, 100.0)
    ws.cell(row=gt_row+3, column=6, value="Overall Waste %").font = Font(bold=True)
    ws.cell(row=gt_row+3, column=6).alignment = Alignment(horizontal='right')
    ws.cell(row=gt_row+3, column=6).border = Border(left=Side(style='thin'), bottom=Side(style='thin'))

    ws.cell(row=gt_row+3, column=7, value=f"{reuse_pct_of_gt:.2f}%").font = Font(bold=True)
    ws.cell(row=gt_row+3, column=7).border = Border(right=Side(style='thin'), bottom=Side(style='thin'))

    # Summary Section (Miscellaneous Cost) - Always show this section
    summary_section_row = gt_row + 5
    ws.cell(row=summary_section_row, column=6, value="MISCELLANEOUS COST").font = Font(bold=True, size=12)
    summary_section_row += 1
    
    # Get summary percentages from project settings file
    summary_pcts = {
        "Overhead Materials": 0.0,
        "Overhead Labor": 0.0,
        "Admin and Management": 0.0,
        "Engineering": 0.0,
        "Packaging Materials": 0.0,
        "Shipping and Transport": 0.0,
        "Commissions": 0.0
    }
    
    # Load percentages from settings file
    print(f"🔍 Attempting to load settings from: {summary_settings_path}")
    print(f"   Elevations path: {elevations_json_path}")
    
    # Always try to construct path from elevations path first (most reliable)
    settings_paths_to_try = []
    
    # 1. Try provided path
    if summary_settings_path:
        settings_paths_to_try.append(summary_settings_path)
        settings_paths_to_try.append(os.path.abspath(summary_settings_path))
    
    # 2. Construct from elevations path (most reliable)
    elev_dir = os.path.dirname(elevations_json_path)
    elev_basename = os.path.basename(elevations_json_path)
    if "_Elevations.json" in elev_basename:
        project_base = elev_basename.replace("_Elevations.json", "")
        constructed_path = os.path.join(elev_dir, f"{project_base}_Settings.json")
        settings_paths_to_try.append(constructed_path)
        print(f"   Constructed from elevations: {constructed_path}")
    
    # 3. Try in same directory as elevations
    if summary_settings_path:
        settings_paths_to_try.append(os.path.join(elev_dir, os.path.basename(summary_settings_path)))
    
    # Remove duplicates while preserving order
    seen = set()
    unique_paths = []
    for path in settings_paths_to_try:
        if path and path not in seen:
            seen.add(path)
            unique_paths.append(path)
    
    settings_loaded = False
    for path_to_try in unique_paths:
        if os.path.exists(path_to_try):
            try:
                print(f"   ✅ Found file, trying to read: {path_to_try}")
                with open(path_to_try, 'r') as f:
                    settings_data = json.load(f)
                    summary_pcts = {
                        "Overhead Materials": settings_data.get("overhead_materials_pct", 0.0),
                        "Overhead Labor": settings_data.get("overhead_labor_pct", 0.0),
                        "Admin and Management": settings_data.get("admin_management_pct", 0.0),
                        "Engineering": settings_data.get("engineering_pct", 0.0),
                        "Packaging Materials": settings_data.get("packaging_materials_pct", 0.0),
                        "Shipping and Transport": settings_data.get("shipping_transport_pct", 0.0),
                        "Commissions": settings_data.get("commissions_pct", 0.0)
                    }
                    print(f"✅ Loaded summary percentages from {path_to_try}")
                    print(f"   Percentages: {summary_pcts}")
                    settings_loaded = True
                    break
            except Exception as e:
                print(f"   ❌ Error reading {path_to_try}: {e}")
                import traceback
                traceback.print_exc()
                continue
    
    if not settings_loaded:
        print(f"⚠️ Could not load settings from any path. Tried:")
        for path_to_try in unique_paths:
            exists = "✅ EXISTS" if os.path.exists(path_to_try) else "❌ NOT FOUND"
            print(f"   {exists}: {path_to_try}")
        if not summary_settings_path:
            print(f"   ⚠️ No summary settings path was provided to create_summary_sheet")
    
    # Calculate base amount: use discounted total only
    base_amount = final_discounted_total
    print(f"📊 Miscellaneous Cost section - Base amount (discounted total): ${base_amount:.2f}")
    
    # Add summary items (only show items with percentages > 0)
    summary_total = 0.0
    items_added = 0
    for label, pct in summary_pcts.items():
        if pct > 0:
            amount = base_amount * (pct / 100.0)
            summary_total += amount
            items_added += 1
            
            ws.cell(row=summary_section_row, column=6, value=label)
            ws.cell(row=summary_section_row, column=7, value=amount).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            summary_section_row += 1
            print(f"   {label}: {pct}% = ${amount:.2f}")
    
    # Add Summary Total if any percentages were set
    if summary_total > 0:
        # Separator line
        ws.cell(row=summary_section_row, column=6).border = Border(top=Side(style='thin'))
        ws.cell(row=summary_section_row, column=7).border = Border(top=Side(style='thin'))
        summary_section_row += 1
        
        ws.cell(row=summary_section_row, column=6, value="MISCELLANEOUS COST TOTAL").font = Font(bold=True)
        ws.cell(row=summary_section_row, column=7, value=summary_total).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=summary_section_row, column=7).font = Font(bold=True)
        summary_section_row += 1
        print(f"✅ Miscellaneous Cost section added with {items_added} items, total: ${summary_total:.2f}")
    else:
        # Show a message if no percentages are set
        ws.cell(row=summary_section_row, column=6, value="(No miscellaneous costs configured)").font = Font(italic=True)
        summary_section_row += 1
        print(f"⚠️ No miscellaneous cost items to add (all percentages are 0)")
        print(f"   Summary settings path was: {summary_settings_path}")
        print(f"   All percentages: {summary_pcts}")
    
    print(f"📝 Miscellaneous Cost section written to rows {gt_row + 5} to {summary_section_row}")

    _autofit_columns(ws, 1, 10, start_row, summary_section_row)
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
    doors=None, mode=None, custom_bay_widths=None, custom_bay_heights=None, summary_settings_path=None
):
    """Generates or updates an Excel report with detailed elevation inputs and calculated outputs."""
    COL_A, COL_B, COL_E, PRICE_COL = 1, 2, 5, 9

    project_root = os.getcwd()
    private_projects_dir = os.path.abspath(os.path.join(project_root, '.files'))
    public_reports_dir = os.path.abspath(os.path.join(project_root, 'reports'))

    os.makedirs(private_projects_dir, exist_ok=True)
    os.makedirs(public_reports_dir, exist_ok=True)
    
    # Normalize paths - use provided path if it exists and is in .files, otherwise construct it
    elevations_path_abs = os.path.abspath(elevations_json_path)
    if os.path.dirname(elevations_path_abs) == private_projects_dir and os.path.exists(elevations_path_abs):
        private_elevations_path = elevations_path_abs
    else:
        private_elevations_path = os.path.join(private_projects_dir, os.path.basename(elevations_json_path))
    
    materials_path_abs = os.path.abspath(extra_materials_json_path)
    if os.path.dirname(materials_path_abs) == private_projects_dir and os.path.exists(materials_path_abs):
        private_extra_materials_path = materials_path_abs
    else:
        private_extra_materials_path = os.path.join(private_projects_dir, os.path.basename(extra_materials_json_path))
    
    excel_path_abs = os.path.abspath(excel_path)
    if os.path.dirname(excel_path_abs) == private_projects_dir:
        private_excel_path = excel_path_abs
    else:
        private_excel_path = os.path.join(private_projects_dir, os.path.basename(excel_path))
    
    # Debug logging
    print(f"📁 Using paths:")
    print(f"   Elevations: {private_elevations_path}")
    print(f"   Materials: {private_extra_materials_path}")
    print(f"   Excel: {private_excel_path}")
    
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
        door_items = calculate_door_info(doors, finish_input, total_count) if doors else []
        
        # Avoid mutating the incoming list reference
        elevation_outputs = base_outputs + door_items

        # Preserve column display preferences from old elevation data if they exist
        old_show_qty_per_elev = False
        old_show_total_cost_per_elev = False
        old_show_discounted_cost_per_elev = False
        if elevation_type in current_saved_elevations:
            old_show_qty_per_elev = current_saved_elevations[elevation_type].get('show_qty_per_elevation', False)
            old_show_total_cost_per_elev = current_saved_elevations[elevation_type].get('show_total_cost_per_elevation', False)
            old_show_discounted_cost_per_elev = current_saved_elevations[elevation_type].get('show_discounted_cost_per_elevation', False)

        current_saved_elevations[elevation_type] = {
            "system": system_input, "finish": finish_input, "total_count": total_count,
            "bays_wide": bays_wide, "bays_tall": bays_tall, "opening_width_inches": opening_width,
            "opening_height_inches": opening_height, "sqft_per_type": sqft_per_type, "total_sqft": total_sqft,
            "perimeter_ft": perimeter_ft, "total_perimeter_ft": total_perimeter_ft,
            "calculated_outputs": elevation_outputs,
            "material_impact": [],
            "custom_bay_widths": custom_bay_widths or [],
            "custom_bay_heights": custom_bay_heights or [],
            "show_qty_per_elevation": old_show_qty_per_elev,
            "show_total_cost_per_elevation": old_show_total_cost_per_elev,
            "show_discounted_cost_per_elevation": old_show_discounted_cost_per_elev
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
            
            # Create a fresh extra_materials state for this elevation (no leftovers from other elevations)
            # This ensures each elevation is calculated independently
            elevation_extra_materials_state = {}

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
            
            # Add bay distribution diagram if custom bays are used
            if elev_data.get("bays_wide") and elev_data.get("bays_tall"):
                diagram_row = current_excel_row + len(input_data) + 3  # Moved up (reduced from 6 to 3)
                custom_widths = elev_data.get('custom_bay_widths', [])
                custom_heights = elev_data.get('custom_bay_heights', [])
                
                # Add diagram label - with 1 more row separation before the picture
                label_cell = ws.cell(row=diagram_row - 2, column=COL_A, value="Bay Distribution Diagram")
                label_cell.font = Font(bold=True, size=12)
                # Picture will be at diagram_row, so there's now 2 rows between label and picture
                
                # Add note in column B (next to bay diagram)
                note_cell = ws.cell(row=diagram_row - 2, column=COL_B, value="*Note - C/L Dimensions")
                note_cell.font = Font(size=12)
                note_cell.alignment = Alignment(horizontal='left', vertical='top')
                
                # Set column widths to accommodate diagram (A-C) without overlapping column D
                ws.column_dimensions['A'].width = 20
                ws.column_dimensions['B'].width = 20
                ws.column_dimensions['C'].width = 20
                
                # Add the diagram
                _add_bay_diagram_to_excel(
                    ws, 
                    diagram_row,
                    elev_data.get("bays_wide", 0),
                    elev_data.get("bays_tall", 0),
                    elev_data.get('opening_width_inches', 0),
                    elev_data.get('opening_height_inches', 0),
                    custom_widths if custom_widths else None,
                    custom_heights if custom_heights else None
                )
            
            output_section_current_row = 1
            profiles_for_section, accessories_for_section, gaskets_for_section, doors_for_section, other_items_for_section = [], [], [], [], []

            current_elevation_finish = elev_data.get("finish")

            # First pass: collect all items and combine same part numbers for profiles/gaskets
            items_by_part = {}  # Key: (part_number, finish), Value: list of items
            
            for item in elev_data.get('calculated_outputs', []):
                pn, manual = item.get('part_number'), item.get('manual', False)
                desc = item.get('description', '').strip()
                is_gasket = "gasket" in desc.lower() or pn in ["E2-0052", "E2-0053", "E2-0065"]
                is_profile = pn in PART_NUMBER_MAP.get("profiles", {})
                is_door = item.get('type', '').lower() in ['door', 'doors']
                
                # For profiles and gaskets, combine items with same part number
                if (is_profile or is_gasket) and pn and pn != "N/A" and not manual:
                    key = (pn, current_elevation_finish)
                    if key not in items_by_part:
                        items_by_part[key] = []
                    items_by_part[key].append(item)
                else:
                    # For other items, add directly to appropriate section
                    if pn and pn != "N/A":
                        if manual and is_door:
                            doors_for_section.append(item)
                        elif manual:
                            other_items_for_section.append(item)
                        elif is_profile:
                            profiles_for_section.append(item)
                        elif is_gasket:
                            gaskets_for_section.append(item)
                        elif pn in PART_NUMBER_MAP.get("accessories", {}) or item.get('type', '').lower() == 'accessory':
                            accessories_for_section.append(item)
                        else:
                            other_items_for_section.append(item)
                    else:
                        if is_door:
                            doors_for_section.append(item)
                        else:
                            other_items_for_section.append(item)
            
            # Combine items with same part number for profiles/gaskets
            for (pn, finish), items_list in items_by_part.items():
                if len(items_list) > 1:
                    # Combine quantities into one list
                    combined_quantities = []
                    combined_descriptions = []
                    total_qty_sum = 0.0
                    total_cost = 0.0
                    
                    for item in items_list:
                        qty = item.get('quantity', 0)
                        if isinstance(qty, list):
                            combined_quantities.extend(qty)
                            qty_sum = sum(qty)
                        else:
                            combined_quantities.append(qty)
                            qty_sum = qty
                        
                        combined_descriptions.append(item.get('description', ''))
                        total_qty_sum += qty_sum
                        # Price is per unit, so multiply by quantity
                        total_cost += item.get('price', 0.0) * qty_sum
                    
                    # Create combined item with weighted average price
                    combined_item = {
                        'description': ' / '.join(combined_descriptions) if len(set(combined_descriptions)) > 1 else combined_descriptions[0],
                        'quantity': combined_quantities,
                        'part_number': pn,
                        'type': items_list[0].get('type', 'profiles'),
                        'price': total_cost / total_qty_sum if total_qty_sum > 0 else 0.0
                    }
                    
                    print(f"DEBUG: Combined {len(items_list)} items for {pn}: {combined_quantities}")
                    
                    # Add to appropriate section
                    if pn in PART_NUMBER_MAP.get("profiles", {}):
                        profiles_for_section.append(combined_item)
                    elif "gasket" in combined_item['description'].lower() or pn in ["E2-0052", "E2-0053", "E2-0065"]:
                        gaskets_for_section.append(combined_item)
                else:
                    # Single item, add directly
                    item = items_list[0]
                    if pn in PART_NUMBER_MAP.get("profiles", {}):
                        profiles_for_section.append(item)
                    elif "gasket" in item.get('description', '').lower() or pn in ["E2-0052", "E2-0053", "E2-0065"]:
                        gaskets_for_section.append(item)

            system_total_for_this_block = [0.0]
            original_system_total_for_this_block = [0.0]
            newly_calculated_material_impacts_for_this_elevation = []
            
            # Get column display preferences and total_count from elevation data
            show_qty_per_elev = elev_data.get('show_qty_per_elevation', False)
            show_total_cost_per_elev = elev_data.get('show_total_cost_per_elevation', False)
            show_discounted_cost_per_elev = elev_data.get('show_discounted_cost_per_elevation', False)
            elev_total_count = elev_data.get('total_count', 1)
            
            # Track totals row numbers for reading from Excel columns
            profile_totals_row = None
            accessory_totals_row = None
            gasket_totals_row = None
            door_totals_row = None

            next_row_after_profiles, impacts_p, profile_totals = _write_output_section(
                ws, "PROFILES", profiles_for_section, COL_E, current_elevation_finish,
                system_total_for_this_block, original_system_total_for_this_block, output_section_current_row,
                elevation_extra_materials_state, private_extra_materials_path, multiplier,
                show_qty_per_elevation=show_qty_per_elev, total_count=elev_total_count,
                show_total_cost_per_elevation=show_total_cost_per_elev, show_discounted_cost_per_elevation=show_discounted_cost_per_elev
            )
            profile_totals_row = next_row_after_profiles - 1  # Totals row is one before next_row

            next_row_after_accessories, impacts_a, accessory_totals = _write_output_section(
                ws, "ACCESSORIES", accessories_for_section, COL_E, current_elevation_finish,
                system_total_for_this_block, original_system_total_for_this_block, next_row_after_profiles,
                elevation_extra_materials_state, private_extra_materials_path, multiplier,
                show_qty_per_elevation=show_qty_per_elev, total_count=elev_total_count,
                show_total_cost_per_elevation=show_total_cost_per_elev, show_discounted_cost_per_elevation=show_discounted_cost_per_elev
            )
            accessory_totals_row = next_row_after_accessories - 1

            next_row_after_gaskets, impacts_g, gasket_totals = _write_output_section(
                ws, "GASKETS", gaskets_for_section, COL_E, current_elevation_finish,
                system_total_for_this_block, original_system_total_for_this_block, next_row_after_accessories,
                elevation_extra_materials_state, private_extra_materials_path, multiplier,
                show_qty_per_elevation=show_qty_per_elev, total_count=elev_total_count,
                show_total_cost_per_elevation=show_total_cost_per_elev, show_discounted_cost_per_elevation=show_discounted_cost_per_elev
            )
            gasket_totals_row = next_row_after_gaskets - 1

            newly_calculated_material_impacts_for_this_elevation.extend(impacts_p)
            newly_calculated_material_impacts_for_this_elevation.extend(impacts_a)
            newly_calculated_material_impacts_for_this_elevation.extend(impacts_g)

            # Process doors section
            next_row_after_doors, impacts_d, door_totals = _write_output_section(
                ws, "DOORS", doors_for_section, COL_E, current_elevation_finish,
                system_total_for_this_block, original_system_total_for_this_block, next_row_after_gaskets,
                elevation_extra_materials_state, private_extra_materials_path, multiplier,
                show_qty_per_elevation=show_qty_per_elev, total_count=elev_total_count,
                show_total_cost_per_elevation=show_total_cost_per_elev, show_discounted_cost_per_elevation=show_discounted_cost_per_elev
            )
            # Only set door_totals_row if there are actually doors
            door_totals_row = (next_row_after_doors - 1) if doors_for_section else None
            newly_calculated_material_impacts_for_this_elevation.extend(impacts_d)

            current_section_row = next_row_after_doors
            grouped_other_misc = {}
            glass_totals_rows = []
            fabrication_totals_rows = []

            for item in other_items_for_section:
                item_type = item.get('type', 'MISCELLANEOUS ITEMS').upper()
                grouped_other_misc.setdefault(item_type, []).append(item)

            for grp_title, grp_items in grouped_other_misc.items():
                next_row_after_group, impacts_g, group_totals = _write_output_section(
                    ws, grp_title, grp_items, COL_E, None,
                    system_total_for_this_block, original_system_total_for_this_block, current_section_row,
                    elevation_extra_materials_state, private_extra_materials_path, multiplier,
                    show_qty_per_elevation=show_qty_per_elev, total_count=elev_total_count,
                    show_total_cost_per_elevation=show_total_cost_per_elev, show_discounted_cost_per_elevation=show_discounted_cost_per_elev
                )
                newly_calculated_material_impacts_for_this_elevation.extend(impacts_g)
                
                # Track totals row for reading from Excel
                group_totals_row = next_row_after_group - 1
                
                # Categorize other items into Glass or Fabrication
                if grp_title == "GLASS" or any(item.get('part_number') == "GLASS_AREA" or item.get('type', '').lower() == 'glass' for item in grp_items):
                    glass_totals_rows.append(group_totals_row)
                else:
                    # Treat other items as fabrication costs
                    fabrication_totals_rows.append(group_totals_row)
                
                current_section_row = next_row_after_group
                print(f"Section '{grp_title}' ended at row {current_section_row}")

            current_saved_elevations[elev_name]['material_impact'] = newly_calculated_material_impacts_for_this_elevation

            # Merge elevation's extra materials state into overall state
            # This accumulates leftovers from all elevations
            print(f"DEBUG: Merging elevation state for {elev_name}, keys: {list(elevation_extra_materials_state.keys())}")
            for key, value in elevation_extra_materials_state.items():
                print(f"DEBUG: Processing key {key}, value: {value}")
                if key not in overall_current_extra_materials_state:
                    overall_current_extra_materials_state[key] = {'quantity': 0, 'length_pieces': []}
                
                # Merge length_pieces (for profiles/gaskets)
                if 'length_pieces' in value and isinstance(value['length_pieces'], list):
                    existing_pieces = overall_current_extra_materials_state[key].get('length_pieces', [])
                    if not isinstance(existing_pieces, list):
                        existing_pieces = []
                    # Combine and sort
                    combined = existing_pieces + value['length_pieces']
                    overall_current_extra_materials_state[key]['length_pieces'] = sorted(combined)
                    print(f"DEBUG: Merged length_pieces for {key}: {combined}")
                
                # Merge quantity (for accessories)
                if 'quantity' in value:
                    overall_current_extra_materials_state[key]['quantity'] = (
                        overall_current_extra_materials_state[key].get('quantity', 0) + value.get('quantity', 0)
                    )
                    print(f"DEBUG: Merged quantity for {key}: {overall_current_extra_materials_state[key]['quantity']}")

            # Create cost breakdown summary table
            # Use the explicitly tracked row position after the last section
            # Add spacing: explicitly create blank rows
            spacing_rows = 1  # Number of blank rows before cost summary (row 39 -> row 43 for headers)
            
            # Explicitly create blank rows by writing empty cells to ensure rows exist
            for blank_row in range(1, spacing_rows + 1):
                # Write empty cell to ensure row exists in Excel
                ws.cell(row=current_section_row + blank_row, column=COL_A, value="")
            
            # Headers start after all spacing rows, with one additional blank row
            cost_summary_row = current_section_row + spacing_rows + 1
            
            # Calculate column numbers for reading from Excel
            # Use the same logic as _write_output_section uses for writing totals row
            # total_col_offset starts at 2 (after Description, Part Number, Total Quantity Required)
            # Then: +1 if qty_per_elev, +1 (label), +1 (Total List Cost), +1 if total_cost_per_elev, then Discounted Total List Cost
            total_col_offset = 2  # Start after "Total Quantity Required"
            if show_qty_per_elev and elev_total_count > 1:
                total_col_offset += 1  # Skip "Quantity Per Elevation" column
            total_col_offset += 1  # Skip to "Total List Cost" column (where label is written)
            total_col_offset += 1  # Skip "Total List Cost" value column
            if show_total_cost_per_elev and elev_total_count > 1:
                total_col_offset += 1  # Skip "Total List Cost Per Elevation" column
            # Now total_col_offset points to "Discounted Total List Cost" column
            
            col_k = COL_E + total_col_offset  # Discounted Total List Cost (Column K)
            col_l = col_k + 1 if (show_discounted_cost_per_elev and elev_total_count > 1) else None  # Discounted Total List Cost Per Elevation (Column L)
            
            # Set cost summary columns - use col_k for total elevation cost column
            header_col = PRICE_COL - 2
            cost_per_elev_col = col_l if col_l else col_k  # Use column L if available, otherwise K
            total_elev_cost_col = col_k  # Use the actual Discounted Total List Cost column
            
            print(f"Column calculation: show_qty_per_elev={show_qty_per_elev}, show_total_cost_per_elev={show_total_cost_per_elev}, show_discounted_cost_per_elev={show_discounted_cost_per_elev}")
            print(f"Total col offset={total_col_offset}, Column K={col_k}, Column L={col_l}")
            
            total_count = elev_data.get("total_count", 1)
            
            # Debug: print row numbers to verify
            print(f"Cost summary for '{elev_name}': Last section row={current_section_row}, Spacing={spacing_rows} rows, Cost summary starts at row={cost_summary_row}")
            print(f"Column K={col_k}, Column L={col_l}")
            
            # Headers on second row
            ws.cell(row=cost_summary_row, column=header_col, value="COST/ELEVATION").font = Font(bold=True)
            ws.cell(row=cost_summary_row, column=cost_per_elev_col, value="COST/ELEVATION").font = Font(bold=True)
            ws.cell(row=cost_summary_row, column=total_elev_cost_col, value="TOTAL ELEVATION COST").font = Font(bold=True)
            
            # Add borders to headers
            for col in [header_col, cost_per_elev_col, total_elev_cost_col]:
                ws.cell(row=cost_summary_row, column=col).border = Border(bottom=Side(style='thin'))
            
            cost_summary_row += 1
            
            # Helper function to read values from Excel cells
            def read_cell_value(row, col):
                """Read cell value, return 0 if None or empty"""
                cell = ws.cell(row=row, column=col)
                return cell.value if cell.value is not None else 0.0
            
            # Profile Costs - read from column L (per elevation) and column K (total)
            profile_cost_per_elev = read_cell_value(profile_totals_row, col_l) if col_l else read_cell_value(profile_totals_row, col_k) / total_count
            profile_total_cost = read_cell_value(profile_totals_row, col_k)
            ws.cell(row=cost_summary_row, column=header_col, value="PROFILE COSTS")
            ws.cell(row=cost_summary_row, column=cost_per_elev_col, value=profile_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            ws.cell(row=cost_summary_row, column=total_elev_cost_col, value=profile_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            cost_summary_row += 1
            
            # Accessory Costs
            accessory_cost_per_elev = read_cell_value(accessory_totals_row, col_l) if col_l else read_cell_value(accessory_totals_row, col_k) / total_count
            accessory_total_cost = read_cell_value(accessory_totals_row, col_k)
            ws.cell(row=cost_summary_row, column=header_col, value="ACCESSORY COSTS")
            ws.cell(row=cost_summary_row, column=cost_per_elev_col, value=accessory_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            ws.cell(row=cost_summary_row, column=total_elev_cost_col, value=accessory_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            cost_summary_row += 1
            
            # Gasket Costs
            gasket_cost_per_elev = read_cell_value(gasket_totals_row, col_l) if col_l else read_cell_value(gasket_totals_row, col_k) / total_count
            gasket_total_cost = read_cell_value(gasket_totals_row, col_k)
            ws.cell(row=cost_summary_row, column=header_col, value="GASKET COSTS")
            ws.cell(row=cost_summary_row, column=cost_per_elev_col, value=gasket_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            ws.cell(row=cost_summary_row, column=total_elev_cost_col, value=gasket_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            cost_summary_row += 1
            
            # Door Costs - only show if there are actually doors
            door_cost_per_elev = read_cell_value(door_totals_row, col_l) if (door_totals_row and col_l) else (read_cell_value(door_totals_row, col_k) / total_count if door_totals_row else 0.0)
            door_total_cost = read_cell_value(door_totals_row, col_k) if door_totals_row else 0.0
            
            # Only display door costs if there are actually doors (total cost > 0)
            if door_totals_row and door_total_cost > 0:
                ws.cell(row=cost_summary_row, column=header_col, value="DOOR COSTS")
                ws.cell(row=cost_summary_row, column=cost_per_elev_col, value=door_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                ws.cell(row=cost_summary_row, column=total_elev_cost_col, value=door_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
                cost_summary_row += 1
            else:
                # No doors, set to 0 for calculations but don't display
                door_cost_per_elev = 0.0
                door_total_cost = 0.0
            
            # Glass Costs - sum from all glass sections
            glass_cost_per_elev = sum(read_cell_value(row, col_l) if col_l else read_cell_value(row, col_k) / total_count for row in glass_totals_rows) if glass_totals_rows else 0.0
            glass_total_cost = sum(read_cell_value(row, col_k) for row in glass_totals_rows) if glass_totals_rows else 0.0
            ws.cell(row=cost_summary_row, column=header_col, value="GLASS COSTS")
            ws.cell(row=cost_summary_row, column=cost_per_elev_col, value=glass_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            ws.cell(row=cost_summary_row, column=total_elev_cost_col, value=glass_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            cost_summary_row += 1
            
            # Fabrication Costs - sum from all fabrication sections
            fabrication_cost_per_elev = sum(read_cell_value(row, col_l) if col_l else read_cell_value(row, col_k) / total_count for row in fabrication_totals_rows) if fabrication_totals_rows else 0.0
            fabrication_total_cost = sum(read_cell_value(row, col_k) for row in fabrication_totals_rows) if fabrication_totals_rows else 0.0
            ws.cell(row=cost_summary_row, column=header_col, value="FABRICATION COSTS")
            ws.cell(row=cost_summary_row, column=cost_per_elev_col, value=fabrication_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            ws.cell(row=cost_summary_row, column=total_elev_cost_col, value=fabrication_total_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            cost_summary_row += 1
            
            # Separator line
            for col in [header_col, cost_per_elev_col, total_elev_cost_col]:
                ws.cell(row=cost_summary_row, column=col).border = Border(top=Side(style='thin'))
            cost_summary_row += 1
            
            # Total Costs - sum from column L (per elevation) and column K (total)
            # Use door costs only if doors exist (door_totals_row is not None and door_total_cost > 0)
            door_cost_for_total = door_cost_per_elev if (door_totals_row and door_total_cost > 0) else 0.0
            door_total_for_total = door_total_cost if (door_totals_row and door_total_cost > 0) else 0.0
            
            total_cost_per_elev = (profile_cost_per_elev + accessory_cost_per_elev + gasket_cost_per_elev + 
                                   door_cost_for_total + glass_cost_per_elev + fabrication_cost_per_elev)
            total_elevation_cost = (profile_total_cost + accessory_total_cost + gasket_total_cost + 
                                   door_total_for_total + glass_total_cost + fabrication_total_cost)
            
            ws.cell(row=cost_summary_row, column=header_col, value=f"{elev_name} TOTAL COSTS").font = Font(bold=True)
            ws.cell(row=cost_summary_row, column=cost_per_elev_col, value=total_cost_per_elev).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            ws.cell(row=cost_summary_row, column=cost_per_elev_col).font = Font(bold=True)
            ws.cell(row=cost_summary_row, column=total_elev_cost_col, value=total_elevation_cost).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            ws.cell(row=cost_summary_row, column=total_elev_cost_col).font = Font(bold=True)
            cost_summary_row += 1
            
            # Note
            note_cell = ws.cell(row=cost_summary_row, column=header_col, value="*Note - Elevation costs based on discounted material costs")
            note_cell.font = Font(italic=True, size=10)
            
            print(f"Rebuilt System Total for '{elev_name}': ${system_total_for_this_block[0]:.2f}")

            # Calculate maximum column based on which optional columns are enabled
            # Column structure: Description (5), Part Number (6), Total Quantity Required (7),
            # [Quantity Per Elevation (8) - optional], Total List Cost (9),
            # [Total List Cost Per Elevation (10) - optional], Discounted Total List Cost (11),
            # [Discounted Total List Cost Per Elevation (12) - optional]
            # Base last column is 11 (Discounted Total List Cost)
            max_col = 11  # Start with base last column (Discounted Total List Cost)
            if show_total_cost_per_elev and elev_total_count > 1:
                max_col += 1  # Total List Cost Per Elevation (column 10)
            if show_discounted_cost_per_elev and elev_total_count > 1:
                max_col += 1  # Discounted Total List Cost Per Elevation (column 12)
            
            _autofit_columns(ws, COL_A, max_col, 1, ws.max_row)
            _clean_trailing_blank_rows(ws, 1)

    save_extra_materials(overall_current_extra_materials_state, private_extra_materials_path)

    summary_ws = wb.create_sheet(title="Summary")
    # Get summary settings path - always construct from elevations path to ensure consistency
    private_summary_settings_path = None
    if summary_settings_path:
        # Extract project base name from elevations path
        elev_basename = os.path.basename(private_elevations_path)
        if "_Elevations.json" in elev_basename:
            project_base = elev_basename.replace("_Elevations.json", "")
            # Construct settings path in the same directory as elevations
            private_summary_settings_path = os.path.join(private_projects_dir, f"{project_base}_Settings.json")
            print(f"🔍 Constructed settings path from elevations: {private_summary_settings_path}")
        else:
            # Fallback to provided path
            summary_settings_path_abs = os.path.abspath(summary_settings_path)
            if os.path.exists(summary_settings_path_abs):
                private_summary_settings_path = summary_settings_path_abs
                print(f"🔍 Using provided settings path: {private_summary_settings_path}")
            else:
                # Try in private projects dir
                private_summary_settings_path = os.path.join(private_projects_dir, os.path.basename(summary_settings_path))
                print(f"🔍 Trying constructed path: {private_summary_settings_path}")
    else:
        # Try to construct from elevations path even if not provided
        elev_basename = os.path.basename(private_elevations_path)
        if "_Elevations.json" in elev_basename:
            project_base = elev_basename.replace("_Elevations.json", "")
            private_summary_settings_path = os.path.join(private_projects_dir, f"{project_base}_Settings.json")
            print(f"🔍 No settings path provided, constructing from elevations: {private_summary_settings_path}")
    
    if private_summary_settings_path and os.path.exists(private_summary_settings_path):
        print(f"✅ Settings file found: {private_summary_settings_path}")
    elif private_summary_settings_path:
        print(f"⚠️ Settings file not found: {private_summary_settings_path}")
    
    create_summary_sheet(summary_ws, private_elevations_path, private_extra_materials_path, wb, summary_settings_path=private_summary_settings_path)
    
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