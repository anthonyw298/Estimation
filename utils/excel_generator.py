import openpyxl
from openpyxl import load_workbook
import os

from data.part_number import PART_NUMBER_MAP
from utils.pricing import get_price_by_part

def generate_excel_report(
    system_input, finish_input, elevation_type, total_count,
    bays_wide, bays_tall, opening_width, opening_height,
    sqft_per_type, total_sqft, perimeter_ft, total_perimeter_ft,
    calculated_outputs, completion_callback=None, mode=None
):
    multiplier = {"clear":1.0, "black":1.1, "paint":1.2}.get(finish_input.lower(), 1.0)
    output_file = "output.xlsx"
    colA, colB, colE = 1, 2, 5

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

    if mode == "new" or not os.path.exists(output_file):
        wb = openpyxl.Workbook()
        ws = wb.active
        start_row = 1
    else:
        wb = load_workbook(output_file)
        ws = wb.active
        running_gt_row = next(
            (r for r in range(1, ws.max_row + 1) if ws.cell(r, colE).value == "RUNNING GRAND TOTAL"),
            None
        )
        start_row = running_gt_row + 3 if running_gt_row else ws.max_row + 2

    # Write inputs vertically
    for i, (h, v) in enumerate(zip(headers, values)):
        ws.cell(row=start_row + i, column=colA, value=h)
        ws.cell(row=start_row + i, column=colB, value=v)

    # Categorize outputs
    profiles, accessories, manual_outputs = [], [], []
    for item in calculated_outputs:
        pn = item.get('part_number')
        typ = item.get('type', '').lower()
        if not pn or pn == "N/A":
            (profiles if typ == "profiles" else accessories if typ == "accessories" else manual_outputs).append(item)
        else:
            if pn in PART_NUMBER_MAP.get("profiles", []):
                profiles.append(item)
            elif pn in PART_NUMBER_MAP.get("accessories", []):
                accessories.append(item)
            else:
                manual_outputs.append(item)

    system_total = 0.0

    def write_section(title, items, col_start, row_start, section_type):
        nonlocal system_total
        headers = ["Description", "Part Number", "Quantity", "Price"]
        for i, h in enumerate([title] + headers):
            ws.cell(row=row_start + i if i > 0 else row_start, column=col_start, value=h)

        write_col = col_start + 1
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

            ws.cell(row=row_start, column=write_col, value=item.get('description', ''))
            ws.cell(row=row_start + 1, column=write_col, value=pn or 'N/A')
            ws.cell(row=row_start + 2, column=write_col, value=f"{qty} {unit_type}")
            ws.cell(row=row_start + 3, column=write_col, value=f"${total_price:.2f}")
            write_col += 1

    cur_row = start_row
    cur_col = colE

    if profiles:
        write_section("PROFILES", profiles, cur_col, cur_row, "profiles")
        cur_row += 6
    if accessories:
        write_section("ACCESSORIES", accessories, cur_col, cur_row, "accessories")
        cur_row += 6
    if manual_outputs:
        grouped = {}
        for item in manual_outputs:
            grouped.setdefault(item.get('type', 'MANUAL').upper(), []).append(item)
        for grp_title, grp_items in grouped.items():
            write_section(grp_title, grp_items, cur_col, cur_row, "manual")
            cur_row += 6

    ws.cell(row=cur_row + 1, column=colE, value="SYSTEM TOTAL")
    ws.cell(row=cur_row + 2, column=colE, value=f"${system_total:.2f}")

    # Remove old running grand totals
    for r in range(ws.max_row, 0, -1):
        if ws.cell(row=r, column=colE).value == "RUNNING GRAND TOTAL":
            ws.delete_rows(r, 2)

    running_grand_total = 0.0
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=colE).value == "SYSTEM TOTAL":
            val = ws.cell(row=r + 1, column=colE).value
            if isinstance(val, str) and val.startswith("$"):
                running_grand_total += float(val.strip("$"))

    bottom_row = ws.max_row + 3
    ws.cell(row=bottom_row, column=colE, value="RUNNING GRAND TOTAL")
    ws.cell(row=bottom_row + 1, column=colE, value=f"${running_grand_total:.2f}")

    # Auto-fit columns
    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value), default=0)
        ws.column_dimensions[col_cells[0].column_letter].width = max_len + 2

    wb.save(output_file)

    if completion_callback:
        completion_callback(
            "Created new output file!" if mode == "new" else "Appended to existing file!",
            "green"
        )
