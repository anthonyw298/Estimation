from openpyxl import load_workbook, Workbook
import os

from data.part_number import PART_NUMBER_MAP
from utils.pricing import get_price_by_part

def generate_excel_report(
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, mode=None, reset=False, all_elevations=None,
    delete_elevation_type=None
):
    multiplier = {"clear": 1.0, "black": 1.1, "paint": 1.2}.get(finish_input.lower(), 1.0)
    output_file = "output.xlsx"
    colA, colB, colE = 1, 2, 5  # Excel columns A=1, B=2, E=5

    headers = [
        "System Input", "Elevation Type", "Total Count",
        "# Bays Wide", "# Bays Tall", "Opening Width", "Opening Height",
        "Sq Ft per Type", "Total Sq Ft", "Perimeter Ft", "Total Perimeter Ft"
    ]
    values = [
        system_input, elevation_type, total_count,
        bays_wide, bays_tall, opening_width, opening_height,
        sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft
    ]

    # Load existing workbook or create new one
    if reset or mode == "new" or not os.path.exists(output_file):
        wb = Workbook()
        ws = wb.active
        start_row = 1
    else:
        wb = load_workbook(output_file)
        ws = wb.active

    def delete_elevation(ws, elevation_name):
        price_col = colE + 3  # Adjust this to your price column index
        delete_start = None
        delete_end = None

        # Find header row with Elevation Type == elevation_name
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=colA).value == "Elevation Type" and ws.cell(row=r, column=colB).value == elevation_name:
                delete_start = r - 1  # Header row before Elevation Type row
                break

        if delete_start is None:
            print(f"Elevation {elevation_name} not found.")
            return

        # Find SYSTEM TOTAL row after delete_start
        for r in range(delete_start + 1, ws.max_row + 1):
            if ws.cell(row=r, column=price_col).value == "SYSTEM TOTAL":
                delete_end = r + 1  # SYSTEM TOTAL row + its price row
                break

        if delete_end is None:
            delete_end = delete_start + 2  # Minimal deletion if SYSTEM TOTAL not found

        # Delete the block of rows for this elevation
        rows_to_delete = delete_end - delete_start + 1
        ws.delete_rows(delete_start, rows_to_delete)

        # Remove existing RUNNING GRAND TOTAL rows anywhere in the sheet
        for r in range(ws.max_row, 0, -1):
            if ws.cell(row=r, column=price_col).value == "RUNNING GRAND TOTAL":
                ws.delete_rows(r, 2)

        # Sum all remaining SYSTEM TOTAL values
        running_grand_total = 0.0
        last_system_total_row = None
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=price_col).value == "SYSTEM TOTAL":
                last_system_total_row = r
                val_cell = ws.cell(row=r + 1, column=price_col).value
                if isinstance(val_cell, str) and val_cell.startswith("$"):
                    try:
                        val = float(val_cell.strip("$"))
                        running_grand_total += val
                    except ValueError:
                        pass

        # Place RUNNING GRAND TOTAL exactly 2 rows below the last SYSTEM TOTAL row
        if last_system_total_row is not None:
            grand_total_label_row = last_system_total_row + 3
            grand_total_value_row = last_system_total_row + 4
        else:
            # If no SYSTEM TOTAL found, append at bottom
            grand_total_label_row = ws.max_row + 1
            grand_total_value_row = ws.max_row + 2

        ws.cell(row=grand_total_label_row, column=price_col, value="RUNNING GRAND TOTAL")
        ws.cell(row=grand_total_value_row, column=price_col, value=f"${running_grand_total:.2f}")

            # Update or insert RUNNING GRAND TOTAL
        found_rgt = False
        for r in range(1, ws.max_row + 1):
                if ws.cell(row=r, column=price_col).value == "RUNNING GRAND TOTAL":
                    ws.cell(row=r + 1, column=price_col).value = f"${running_grand_total:.2f}"
                    found_rgt = True
                    break

        if not found_rgt:
                # Add RUNNING GRAND TOTAL at bottom
                last_row = ws.max_row + 2
                ws.cell(row=last_row, column=price_col, value="RUNNING GRAND TOTAL")
                ws.cell(row=last_row + 1, column=price_col, value=f"${running_grand_total:.2f}")

    # If deletion requested, perform it and exit
    if delete_elevation_type:
        delete_elevation(ws, delete_elevation_type)
        wb.save(output_file)
        if completion_callback:
            completion_callback()
        return

    # Determine start row for writing new data
    has_elevations = False
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=colA).value == "Elevation Type":
            has_elevations = True
            break

    if reset or not has_elevations:
        ws.delete_rows(1, ws.max_row)
        start_row = 1
    else:
        start_row = None

        # Try to find existing elevation header
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=colA).value == "Elevation Type" and ws.cell(row=r, column=colB).value == elevation_type:
                start_row = r - 1
                break
        # If not found, start after RUNNING GRAND TOTAL or at bottom
        if start_row is None:
            running_gt_row = next(
                (r for r in range(1, ws.max_row + 1) if ws.cell(r, colE + 3).value == "RUNNING GRAND TOTAL"),
                None
            )
            if running_gt_row:
                start_row = running_gt_row + 3
            else:
                start_row = ws.max_row + 2 if ws.max_row > 1 or ws.cell(1, colA).value else 1

    # Write main headers and values if not resetting or no all_elevations flag
    if not (reset or not all_elevations):
        for i, (header, val) in enumerate(zip(headers, values)):
            ws.cell(row=start_row + i, column=colA, value=header)
            ws.cell(row=start_row + i, column=colB, value=val)

    system_total = 0.0

    if all_elevations and len(all_elevations) > 0 and not reset:
        profiles, accessories, manual_outputs = [], [], []

        # Categorize calculated outputs into profiles, accessories, and manual
        for item in calculated_outputs:
            pn = item.get('part_number')
            typ = item.get('type', '').lower()
            if not pn or pn == "N/A":
                if typ == "profiles":
                    profiles.append(item)
                elif typ == "accessories":
                    accessories.append(item)
                else:
                    manual_outputs.append(item)
            else:
                if pn in PART_NUMBER_MAP.get("profiles", []):
                    profiles.append(item)
                elif pn in PART_NUMBER_MAP.get("accessories", []):
                    accessories.append(item)
                else:
                    manual_outputs.append(item)

        def write_section(title, items, col_start, row_start, section_type):
            nonlocal system_total
            headers = ["Description", "Part Number", "Quantity", "Price"]

            ws.cell(row=row_start, column=col_start, value=title)
            for i, h in enumerate(headers):
                ws.cell(row=row_start + 1, column=col_start + i, value=h)

            write_row = row_start + 2
            for item in items:
                qty = item.get('quantity', 0)
                pn = item.get('part_number')
                if pn and pn != "N/A":
                    unit_price, unit_type = get_price_by_part(pn)
                    unit_price = unit_price or 0.0
                    unit_type = unit_type or "pcs"
                else:
                    unit_price = item.get('price', 0.0)
                    unit_type = item.get('unit', 'pcs')

                if section_type == "profiles":
                    unit_price *= multiplier

                total_price = qty * unit_price
                system_total += total_price

                ws.cell(row=write_row, column=col_start, value=item.get('description', ''))
                ws.cell(row=write_row, column=col_start + 1, value=pn or 'N/A')
                ws.cell(row=write_row, column=col_start + 2, value=f"{qty} {unit_type}")
                ws.cell(row=write_row, column=col_start + 3, value=f"${total_price:.2f}")

                write_row += 1

            return write_row + 1

        cur_row = start_row
        cur_col = colE

        if profiles:
            cur_row = write_section("PROFILES", profiles, cur_col, cur_row, "profiles")
        if accessories:
            cur_row = write_section("ACCESSORIES", accessories, cur_col, cur_row, "accessories")
        if manual_outputs:
            grouped = {}
            for item in manual_outputs:
                grouped.setdefault(item.get('type', 'MANUAL').upper(), []).append(item)
            for grp_title, grp_items in grouped.items():
                cur_row = write_section(grp_title, grp_items, cur_col, cur_row, "manual")

        ws.cell(row=cur_row + 1, column=colE + 3, value="SYSTEM TOTAL")
        ws.cell(row=cur_row + 2, column=colE + 3, value=f"${system_total:.2f}")

        # Remove old RUNNING GRAND TOTAL rows if exist
        for r in range(ws.max_row, 0, -1):
            if ws.cell(row=r, column=colE + 3).value == "RUNNING GRAND TOTAL":
                ws.delete_rows(r, 3)

        # Calculate new RUNNING GRAND TOTAL by summing all SYSTEM TOTAL values
        running_grand_total = 0.0
        for r in range(1, ws.max_row):
            if ws.cell(row=r, column=colE + 3).value == "SYSTEM TOTAL":
                val_cell = ws.cell(row=r + 1, column=colE + 3).value
                if isinstance(val_cell, str) and val_cell.startswith("$"):
                    try:
                        val = float(val_cell.strip("$"))
                        running_grand_total += val
                    except ValueError:
                        pass

        last_row = ws.max_row + 2
        ws.cell(row=last_row, column=colE + 3, value="RUNNING GRAND TOTAL")
        ws.cell(row=last_row + 1, column=colE + 3, value=f"${running_grand_total:.2f}")

    # Adjust column widths automatically for readability
    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value), default=0)
        ws.column_dimensions[col_cells[0].column_letter].width = max_len + 2

    wb.save(output_file)
    if completion_callback:
        completion_callback()
