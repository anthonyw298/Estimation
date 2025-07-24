import os
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, numbers
from openpyxl.utils import get_column_letter

from data.part_number import PART_NUMBER_MAP
from utils.pricing import get_price_by_part

output_file = "output.xlsx"


def create_summary_sheet(excel_path=output_file, json_path='saved_elevations.json'):
    """
    Reads saved elevation data, aggregates quantities by part number,
    calculates total prices, and writes a summary section into the Excel file.
    """
    try:
        with open(json_path, 'r') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"⚠️ JSON file '{json_path}' not found. Skipping summary sheet creation.")
        return

    try:
        wb = load_workbook(excel_path)
    except FileNotFoundError:
        print(f"⚠️ Excel file '{excel_path}' not found. Cannot update summary sheet.")
        return

    ws = wb.active

    # Check if summary exists by scanning column A for "Part Number"
    summary_start_row = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val == "Part Number":
            summary_start_row = r
            break

    # If summary exists and data length <= 1, delete all summary rows until a blank row
    if summary_start_row and len(data) <= 1:
        current_row = summary_start_row
        # Keep deleting rows while column A is not empty
        while current_row <= ws.max_row and ws.cell(row=current_row, column=1).value not in (None, ""):
            ws.delete_rows(current_row, 1)
            # Don't increment current_row here because rows shift up after deletion
        wb.save(excel_path)
        print("ℹ️ Only one or zero elevations found. Existing Summary sheet deleted.")
        return
    elif len(data) <= 1:
        print("ℹ️ Only one or zero elevations found. No summary created.")
        return

    # Otherwise continue to build summary if more than 1 data entry
    # Aggregate quantities by part_number
    summary = {}
    for elev_data in data.values():
        outputs = elev_data.get('calculated_outputs', [])
        for output in outputs:
            part_number = output.get('part_number')
            quantity = output.get('quantity', 0)
            if part_number:
                summary[part_number] = summary.get(part_number, 0) + quantity

    # Calculate prices and totals
    summary_rows = []
    for part, qty in summary.items():
        price_per_unit, _ = get_price_by_part(part)
        price_per_unit = price_per_unit or 0
        total_price = price_per_unit * qty
        summary_rows.append((part, qty, total_price))

    # Remove old summary section if exists (search for header "Part Number")
    if summary_start_row:
        ws.delete_rows(summary_start_row, ws.max_row - summary_start_row + 1)

    # Find last "System Total" row to insert summary below it
    last_sys_total = None
    for r in range(ws.max_row, 0, -1):
        cell_val = ws.cell(row=r, column=1).value
        if cell_val and "System Total" in str(cell_val):
            last_sys_total = r
            break
    if not last_sys_total:
        last_sys_total = ws.max_row

    start_row = last_sys_total + 2

    # Write headers for summary table
    ws.cell(row=start_row, column=1, value="Part Number").font = Font(bold=True)
    ws.cell(row=start_row, column=2, value="Total Quantity").font = Font(bold=True)
    ws.cell(row=start_row, column=3, value="Total Price").font = Font(bold=True)

    # Write summary rows
    for idx, (part, qty, total) in enumerate(summary_rows, start=start_row + 1):
        ws.cell(row=idx, column=1, value=part)
        ws.cell(row=idx, column=2, value=qty)
        price_cell = ws.cell(row=idx, column=3, value=total)
        price_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

    # Autofit columns (safe check for None values)
    for col in range(1, 4):
        col_letter = get_column_letter(col)
        max_len = 0
        for r in range(start_row, ws.max_row + 1):
            val = ws.cell(row=r, column=col).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = max_len + 2

    wb.save(excel_path)
    print(f"✅ Summary sheet updated in {excel_path}.")

def generate_excel_report(
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, mode=None, reset=False, all_elevations=None,
    delete_elevation_type=None
):
    multiplier = {"clear": 1.0, "black": 1.1, "paint": 1.2}.get(finish_input.lower(), 1.0)
    output_file = "output.xlsx"
    colA, colB, colE = 1, 2, 5  # A=1, B=2, E=5

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

    # Load or create workbook
    if reset or mode == "new" or not os.path.exists(output_file):
        wb = Workbook()
        ws = wb.active
        ws.title = "Report"
    else:
        wb = load_workbook(output_file)
        ws = wb.active

    def delete_elevation(ws, elevation_name):
        price_col = colE + 3  # SYSTEM TOTAL col
        delete_start, delete_end = None, None

        # Find start of elevation
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=colA).value == "Elevation Type" and ws.cell(row=r, column=colB).value == elevation_name:
                delete_start = r - 1
                break

        if delete_start is None:
            print(f"Elevation '{elevation_name}' not found.")
            return

        # Find SYSTEM TOTAL for that elevation
        for r in range(delete_start + 1, ws.max_row + 1):
            if ws.cell(row=r, column=price_col).value == "SYSTEM TOTAL":
                delete_end = r + 1
                break

        if delete_end is None:
            delete_end = delete_start + 2  # fallback

        ws.delete_rows(delete_start, delete_end - delete_start + 1)

        # Remove old RUNNING GRAND TOTALS
        for r in range(ws.max_row, 0, -1):
            if ws.cell(row=r, column=price_col).value == "RUNNING GRAND TOTAL":
                ws.delete_rows(r, 2)

        # Recalculate & rewrite new RUNNING GRAND TOTAL
        running_grand_total = 0.0
        last_system_total_row = None
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=price_col).value == "SYSTEM TOTAL":
                last_system_total_row = r
                val = ws.cell(row=r + 1, column=price_col).value
                if isinstance(val, str) and val.startswith("$"):
                    try:
                        running_grand_total += float(val.strip("$"))
                    except:
                        pass

        if last_system_total_row:
            new_gt_row = last_system_total_row + 2
        else:
            new_gt_row = ws.max_row + 1

        ws.cell(row=new_gt_row + 1, column=price_col, value="RUNNING GRAND TOTAL")
        ws.cell(row=new_gt_row + 2, column=price_col, value=f"${running_grand_total:.2f}")

    # --- If in delete mode ---
    if delete_elevation_type:
        delete_elevation(ws, delete_elevation_type)
        wb.save(output_file)
        create_summary_sheet(excel_path=output_file)
        if completion_callback:
            completion_callback()
        return

    # If not deleting: also auto delete this elevation block to prevent duplicates
    if not reset and elevation_type:
        delete_elevation(ws, elevation_type)

    # Find insertion row
    has_elevations = any(
        ws.cell(row=r, column=colA).value == "Elevation Type"
        for r in range(1, ws.max_row + 1)
    )

    if reset or not has_elevations:
        ws.delete_rows(1, ws.max_row)
        start_row = 1
    else:
        last_gt = next(
            (r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=colE + 3).value == "RUNNING GRAND TOTAL"),
            None
        )
        start_row = last_gt + 3 if last_gt else ws.max_row + 2

    # Write headers/inputs
    if not (reset or not all_elevations):
        for i, (header, val) in enumerate(zip(headers, values)):
            ws.cell(row=start_row + i, column=colA, value=header)
            ws.cell(row=start_row + i, column=colB, value=val)

    # Write calculated
    system_total = 0.0

    if all_elevations and not reset:
        profiles, accessories, manual_outputs = [], [], []

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

        def write_section(title, items, row_start):
            nonlocal system_total
            headers = ["Description", "Part Number", "Quantity", "Price"]
            ws.cell(row=row_start, column=colE, value=title)
            for i, h in enumerate(headers):
                ws.cell(row=row_start + 1, column=colE + i, value=h)
            cur_row = row_start 
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

                if title == "PROFILES":
                    unit_price *= multiplier

                total_price = qty * unit_price
                system_total += total_price

                ws.cell(row=cur_row, column=colE, value=item.get('description', ''))
                ws.cell(row=cur_row, column=colE + 1, value=pn or 'N/A')
                ws.cell(row=cur_row, column=colE + 2, value=f"{qty} {unit_type}")
                ws.cell(row=cur_row, column=colE + 3, value=f"${total_price:.2f}")
                cur_row += 1

            return cur_row + 1

        cur_row = start_row
        if profiles:
            cur_row = write_section("PROFILES", profiles, cur_row)
        if accessories:
            cur_row = write_section("ACCESSORIES", accessories, cur_row)
        if manual_outputs:
            grouped = {}
            for item in manual_outputs:
                grouped.setdefault(item.get('type', 'MANUAL').upper(), []).append(item)
            for grp_title, grp_items in grouped.items():
                cur_row = write_section(grp_title, grp_items, cur_row)

        # SYSTEM TOTAL
        ws.cell(row=cur_row, column=colE + 3, value="SYSTEM TOTAL")
        ws.cell(row=cur_row + 1, column=colE + 3, value=f"${system_total:.2f}")

        # Remove old RUNNING GRAND TOTALS
        for r in range(ws.max_row, 0, -1):
            if ws.cell(row=r, column=colE + 3).value == "RUNNING GRAND TOTAL":
                ws.delete_rows(r, 2)

        # Recompute RUNNING GRAND TOTAL
        running_grand_total = 0.0
        last_system_total_row = None
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=colE + 3).value == "SYSTEM TOTAL":
                last_system_total_row = r
                val = ws.cell(row=r + 1, column=colE + 3).value
                if isinstance(val, str) and val.startswith("$"):
                    try:
                        running_grand_total += float(val.strip("$"))
                    except:
                        pass

        if last_system_total_row:
            new_gt_row = last_system_total_row + 2
        else:
            new_gt_row = ws.max_row + 1

        ws.cell(row=new_gt_row + 1, column=colE + 3, value="RUNNING GRAND TOTAL")
        ws.cell(row=new_gt_row + 2, column=colE + 3, value=f"${running_grand_total:.2f}")

    # Autosize
    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value), default=0)
        ws.column_dimensions[col_cells[0].column_letter].width = max_len + 2

    wb.save(output_file)
    create_summary_sheet(excel_path=output_file)
    if completion_callback:
        completion_callback()
