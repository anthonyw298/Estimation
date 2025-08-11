def _write_output_section(ws, title, items, colE, elevation_finish, system_total_ref, start_output_row, current_extra_materials_state, extra_materials_path, multiplier):
    """Writes a section of calculated outputs to the worksheet, applying the discount multiplier."""
    if not items: return start_output_row, []

    current_row = start_output_row
    ws.cell(row=current_row, column=colE, value=title).font = Font(bold=True)
    # Add a new column for "Discounted Price"
    for i, h in enumerate(["Description", "Part Number", "Quantity", "Original Price", "Discounted Price"]):
        ws.cell(row=current_row + 1, column=colE + i, value=h).font = Font(bold=True)
    current_row += 2

    section_material_impacts = []
    
    # Calculate the total for this section
    section_original_total = 0.0

    for item in items:
        qty_raw = item.get('quantity', 0)
        pn, manual = item.get('part_number'), item.get('manual', False)

        individual_quantities = qty_raw if isinstance(qty_raw, list) else [qty_raw]

        if isinstance(qty_raw, list):
            if len(qty_raw) > 1 and all(x == qty_raw[0] for x in qty_raw):
                display_qty_string = f"{qty_raw[0]:.2f} x {len(qty_raw)}"
            else:
                display_qty_string = ", ".join([f"{q:.2f}" for q in qty_raw])
        else:
            display_qty_string = f"{qty_raw:.2f}"

        item_total_original_price = 0.0

        for single_qty_for_calc in individual_quantities:
            total_item_price_single_cut, unit_type, material_impact_details = 0.0, "pcs", None

            # Get the price from the stored JSON if available, otherwise calculate it
            original_price = item.get('original_price')

            if original_price is not None:
                # If original price is already in the JSON, use it
                total_item_price_single_cut = original_price * single_qty_for_calc
                unit_type = item.get('unit', 'pcs')
            else:
                # Fallback to get_price_by_part if price is not stored
                if manual:
                    if pn and pn != "N/A":
                        price_calculated, unit_calculated, material_impact_details = \
                            get_price_by_part(pn, single_qty_for_calc, finish=elevation_finish, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False, group=True)  
                        total_item_price_single_cut = (price_calculated if price_calculated is not None else item.get('price', 0.0) * single_qty_for_calc)
                        unit_type = item.get('unit') or unit_calculated or 'pcs'
                    else:
                        total_item_price_single_cut = item.get('price', 0.0) * single_qty_for_calc
                        unit_type = item.get('unit', 'pcs')
                        material_impact_details = {
                            'part_number': "N/A - Manual", 'requested_qty': single_qty_for_calc, 'purchased_qty_or_length': 0.0,
                            'leftover_generated_qty_or_length': 0.0, 'used_from_leftover_qty_or_length': 0.0,
                            'cost_incurred': total_item_price_single_cut, 'type_processed_as': 'manual_no_pn',
                            'finish': None
                        }
                else:
                    price, unit_type, material_impact_details = \
                        get_price_by_part(pn, single_qty_for_calc, finish=elevation_finish, current_extra_materials=current_extra_materials_state, extra_materials_file=extra_materials_path, summary=False)
                    total_item_price_single_cut = price or 0.0
                    unit_type = unit_type or "pcs"
            
            item_total_original_price += total_item_price_single_cut

            if material_impact_details:
                material_impact_details['leftover_generated_qty_or_length_display'] = f"{material_impact_details.get('leftover_generated_qty_or_length', 0.0):.2f}"
                section_material_impacts.append(material_impact_details)
                apply_material_impact_to_extra_materials_in_memory(current_extra_materials_state, material_impact_details)
        
        # Apply the discount multiplier ONLY to profiles and accessories
        item_type = item.get('type')
        if item_type in ["profiles", "accessories"]:
            discounted_price_for_display = item_total_original_price * multiplier
        else:
            # Other items like glass, labor, etc. are not discounted
            discounted_price_for_display = item_total_original_price
            
        system_total_ref[0] += discounted_price_for_display

        ws.cell(row=current_row, column=colE, value=item.get('description', ''))
        ws.cell(row=current_row, column=colE + 1, value=pn or 'N/A')
        ws.cell(row=current_row, column=colE + 2, value=f"{display_qty_string} {unit_type}")
        ws.cell(row=current_row, column=colE + 3, value=item_total_original_price).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        ws.cell(row=current_row, column=colE + 4, value=discounted_price_for_display).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        current_row += 1
    
    # Add a SYSTEM TOTAL and DISCOUNTED SYSTEM TOTAL for the current elevation
    system_total_row = current_row
    ws.cell(row=system_total_row, column=colE + 3, value="SYSTEM TOTAL").font = Font(bold=True)
    ws.cell(row=system_total_row, column=colE + 4, value=system_total_ref[0]).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
    
    return current_row + 1, section_material_impacts