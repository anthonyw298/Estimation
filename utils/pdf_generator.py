"""
PDF Export functionality for project reports.
Generates PDF directly from report_data dict (no Excel parsing).
"""

import os
import io
from datetime import datetime

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate,
        Table,
        TableStyle,
        Paragraph,
        Spacer,
        Image,
        PageBreak,
        KeepTogether,
    )
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


def get_logo_path():
    """Get the path to the company logo."""
    possible_paths = [
        os.path.join("assets", "logo.png"),
        os.path.join("assets", "logo.jpg"),
        os.path.join("assets", "logo.jpeg"),
        os.path.join("assets", "company_logo.png"),
        os.path.join("assets", "company_logo.jpg"),
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None


def _fmt_currency(val):
    """Format a numeric value as currency string."""
    if val is None:
        return ""
    try:
        v = float(val)
        return f"${v:,.2f}"
    except (ValueError, TypeError):
        return str(val)


def _fmt_pct(val):
    """Format a numeric value as percentage string."""
    if val is None or val == "N/A":
        return "N/A"
    try:
        v = float(val)
        return f"{v:.2f}%"
    except (ValueError, TypeError):
        return str(val)


# ---------------------------------------------------------------------------
# Shared PDF table builder
# ---------------------------------------------------------------------------


def _create_pdf_table(
    table_data, headers, normal_style, section_header=None, section_style=None
):
    """Create a formatted PDF table from data.

    table_data: list of rows. Each row is a list of values.
                Values can be plain strings or (value, is_bold) tuples.
    headers:    list of column header strings.
    """
    if not REPORTLAB_AVAILABLE:
        return []

    elements = []
    if not table_data or not headers:
        return elements

    full_data = []

    # Header row with wrapping
    wrapped_headers = []
    for h in headers:
        if h:
            if len(str(h)) > 20:
                words = str(h).split()
                lines, current = [], ""
                for word in words:
                    test = current + " " + word if current else word
                    if len(test) <= 20:
                        current = test
                    else:
                        if current:
                            lines.append(current)
                        current = word
                if current:
                    lines.append(current)
                wrapped_headers.append(Paragraph("<br/>".join(lines), normal_style))
            else:
                wrapped_headers.append(Paragraph(str(h), normal_style))
        else:
            wrapped_headers.append("")
    full_data.append(wrapped_headers)

    # Data rows
    for row in table_data:
        formatted = []
        for i, cell in enumerate(row):
            if i >= len(headers):
                break
            cell_value = cell
            is_bold = False
            if isinstance(cell, tuple) and len(cell) == 2:
                cell_value, is_bold = cell
            if cell_value is not None:
                cell_str = str(cell_value).strip()
                if cell_str:
                    if is_bold:
                        formatted.append(Paragraph(f"<b>{cell_str}</b>", normal_style))
                    else:
                        formatted.append(Paragraph(cell_str, normal_style))
                else:
                    formatted.append("")
            else:
                formatted.append("")
        while len(formatted) < len(headers):
            formatted.append("")
        full_data.append(formatted[: len(headers)])

    if len(full_data) <= 1:
        return elements

    num_cols = len(headers)
    total_width = 7.2 * inch

    if num_cols <= 3:
        col_widths = [total_width / num_cols] * num_cols
    elif num_cols <= 5:
        desc_width = total_width * 0.3
        other_width = (total_width - desc_width) / (num_cols - 1)
        col_widths = [desc_width] + [other_width] * (num_cols - 1)
    else:
        desc_width = total_width * 0.22
        other_width = (total_width - desc_width) / (num_cols - 1)
        col_widths = [desc_width] + [other_width] * (num_cols - 1)

    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8E8E8")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1A1A1A")),
        ("ALIGN", (0, 0), (-1, 0), "LEFT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
        ("LEFTPADDING", (0, 0), (-1, 0), 4),
        ("RIGHTPADDING", (0, 0), (-1, 0), 4),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("TOPPADDING", (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
        ("LEFTPADDING", (0, 1), (-1, -1), 4),
        ("RIGHTPADDING", (0, 1), (-1, -1), 4),
        ("VALIGN", (0, 1), (-1, -1), "TOP"),
        (
            "ROWBACKGROUNDS",
            (0, 1),
            (-1, -1),
            [colors.white, colors.HexColor("#F9F9F9")],
        ),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
    ]

    tbl = Table(
        full_data, colWidths=col_widths[:num_cols], repeatRows=1, splitByRow=True
    )
    tbl.setStyle(TableStyle(table_style))

    if section_header and section_style:
        keep = [
            Paragraph(f"<b>{section_header}</b>", section_style),
            Spacer(1, 0.1 * inch),
            tbl,
        ]
        elements.append(KeepTogether(keep))
        elements.append(Spacer(1, 0.2 * inch))
    else:
        elements.append(tbl)
        elements.append(Spacer(1, 0.2 * inch))

    return elements


# ---------------------------------------------------------------------------
# Main entry: generate PDF from report_data dict
# ---------------------------------------------------------------------------


def generate_pdf_from_data(report_data, pdf_path, include_logo=True):
    """Generate a PDF report directly from a report_data dict.

    report_data structure:
    {
      "project_name": str,
      "report_options": { ... checkbox state ... },
      "elevations": {
        "<name>": {
          "system_input": [ {"label": str, "value": str}, ... ],
          "sections": {
            "profiles": [ {item_fields...}, ... ],
            "accessories": [...], "gaskets": [...], "doors": [...],
            "glass": [...], "fabrication": [...]
          },
          "section_totals": { "profiles": {"original": float, "discounted": float}, ... },
          "cost_summary": { "profile_cost_per_elev": float, ... },
          "has_doors": bool,
          "total_count": int,
        }
      },
      "summary": {
        "categories": { "PROFILES": [ {item_fields...} ], ... },
        "category_totals": { "PROFILES": {"original": float, "discounted": float, "residual": float}, ... },
        "cost_overview": { "list_price_total": float, "discounted_total": float, "residual_waste_cost": float, "waste_pct": float },
        "additional_costs": { "items": [(label, amount)], "total": float },
        "markups": { "items": [(label, amount)], "total": float },
        "project_total": float,
        "elevation_summary": [ {"name": str, "quantity": int, "dimensions": str, "sqft_total": float, "perimeter_ft_total": float} ],
        "elevation_summary_totals": { "total_qty": float, "total_sqft": float, "total_perimeter": float },
      }
    }
    """
    if not REPORTLAB_AVAILABLE:
        raise ImportError(
            "reportlab is not installed. Install with: pip install reportlab"
        )

    opts = report_data.get("report_options", {})
    elevations_included = opts.get("elevations_included", {})
    summary_included = opts.get("summary_included", True)
    per_elev_sections = opts.get("per_elevation_sections", {})
    per_elev_columns = opts.get("per_elevation_columns", {})
    summary_opts = opts.get("summary_options", {})

    # Create PDF document
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        rightMargin=0.4 * inch,
        leftMargin=0.4 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.5 * inch,
    )

    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Heading1"],
        fontSize=18,
        textColor=colors.HexColor("#1A1A1A"),
        spaceAfter=15,
        alignment=TA_CENTER,
        fontName="Helvetica-Bold",
    )
    section_style = ParagraphStyle(
        "SectionStyle",
        parent=styles["Heading2"],
        fontSize=12,
        textColor=colors.HexColor("#1A1A1A"),
        spaceAfter=8,
        spaceBefore=12,
        fontName="Helvetica-Bold",
    )
    normal_style = ParagraphStyle(
        "CustomNormal",
        parent=styles["Normal"],
        fontSize=8,
        textColor=colors.HexColor("#333333"),
        spaceAfter=3,
        fontName="Helvetica",
    )

    # Logo
    if include_logo:
        logo_path = get_logo_path()
        if logo_path:
            try:
                logo = Image(logo_path, width=2.5 * inch, height=1 * inch)
                logo.hAlign = "CENTER"
                story.append(logo)
                story.append(Spacer(1, 0.15 * inch))
            except Exception:
                pass

    # Title
    project_name = report_data.get("project_name", "Project")
    story.append(Paragraph("PROJECT ESTIMATION REPORT", title_style))
    story.append(Paragraph(f"<b>Project:</b> {project_name}", normal_style))
    story.append(
        Paragraph(f"<b>Date:</b> {datetime.now().strftime('%B %d, %Y')}", normal_style)
    )
    story.append(Spacer(1, 0.3 * inch))

    elev_data_all = report_data.get("elevations", {})

    # =========== ELEVATION SHEETS ===========
    for elev_name, elev_data in elev_data_all.items():
        if not elevations_included.get(elev_name, True):
            continue

        elev_sec = per_elev_sections.get(elev_name, {})
        elev_col = per_elev_columns.get(elev_name, {})
        total_count = elev_data.get("total_count", 1)

        story.append(Paragraph(f"<b>{elev_name}</b>", section_style))
        story.append(Spacer(1, 0.2 * inch))

        # --- System Input ---
        if elev_sec.get("system_input", True):
            si_items = elev_data.get("system_input", [])
            if si_items:
                si_table = [
                    [item.get("label", ""), item.get("value", "")] for item in si_items
                ]
                story.extend(
                    _create_pdf_table(
                        si_table,
                        ["Item", "Value"],
                        normal_style,
                        "System Input",
                        section_style,
                    )
                )

        # --- Material Sections ---
        SECTION_ORDER = [
            "profiles",
            "accessories",
            "gaskets",
            "doors",
            "glass",
            "fabrication",
        ]
        SECTION_TITLES = {
            "profiles": "PROFILES",
            "accessories": "ACCESSORIES",
            "gaskets": "GASKETS",
            "doors": "DOORS",
            "glass": "GLASS",
            "fabrication": "FABRICATION",
        }

        sections_data = elev_data.get("sections", {})
        section_totals = elev_data.get("section_totals", {})

        for sec_key in SECTION_ORDER:
            if not elev_sec.get(sec_key, True):
                continue
            # Skip doors if elevation has no doors
            if sec_key == "doors" and not elev_data.get("has_doors", False):
                continue

            items = sections_data.get(sec_key, [])
            if not items:
                continue

            sec_title = SECTION_TITLES.get(sec_key, sec_key.upper())

            # Build headers based on column config
            headers = []
            col_keys = []  # track which keys map to which column
            if elev_col.get("description", True):
                headers.append("Description")
                col_keys.append("description")
            if elev_col.get("part_number", True):
                headers.append("Part Number")
                col_keys.append("part_number")
            if elev_col.get("total_quantity_required", True):
                headers.append("Total Quantity Required")
                col_keys.append("display_qty")
            if elev_col.get("quantity_per_elevation", True) and total_count > 1:
                headers.append("Qty Per Elevation")
                col_keys.append("qty_per_elev")
            if elev_col.get("total_list_cost", True):
                headers.append("Total List Cost")
                col_keys.append("original_cost")
            if elev_col.get("total_list_cost_per_elevation", True) and total_count > 1:
                headers.append("List Cost Per Elev")
                col_keys.append("original_cost_per_elev")
            if elev_col.get("discounted_total_list_cost", True):
                headers.append("Discounted Total Cost")
                col_keys.append("discounted_cost")
            if (
                elev_col.get("discounted_total_list_cost_per_elevation", True)
                and total_count > 1
            ):
                headers.append("Discounted Per Elev")
                col_keys.append("discounted_cost_per_elev")

            if not headers:
                continue

            # Build data rows
            table_rows = []
            for item in items:
                row = []
                for ck in col_keys:
                    if ck == "description":
                        row.append(item.get("description", ""))
                    elif ck == "part_number":
                        row.append(item.get("part_number", ""))
                    elif ck == "display_qty":
                        row.append(item.get("display_qty", ""))
                    elif ck == "qty_per_elev":
                        row.append(item.get("qty_per_elev", ""))
                    elif ck == "original_cost":
                        row.append(_fmt_currency(item.get("original_cost")))
                    elif ck == "original_cost_per_elev":
                        row.append(_fmt_currency(item.get("original_cost_per_elev")))
                    elif ck == "discounted_cost":
                        row.append(_fmt_currency(item.get("discounted_cost")))
                    elif ck == "discounted_cost_per_elev":
                        row.append(_fmt_currency(item.get("discounted_cost_per_elev")))
                    else:
                        row.append("")
                table_rows.append(row)

            # Add totals row (bold)
            totals = section_totals.get(sec_key, {})
            if totals:
                total_row = []
                for ck in col_keys:
                    if ck == "description":
                        total_row.append((f"Total {sec_title.title()} Cost", True))
                    elif ck == "original_cost":
                        total_row.append((_fmt_currency(totals.get("original")), True))
                    elif ck == "original_cost_per_elev":
                        total_row.append(
                            (_fmt_currency(totals.get("original_per_elev")), True)
                        )
                    elif ck == "discounted_cost":
                        total_row.append(
                            (_fmt_currency(totals.get("discounted")), True)
                        )
                    elif ck == "discounted_cost_per_elev":
                        total_row.append(
                            (_fmt_currency(totals.get("discounted_per_elev")), True)
                        )
                    else:
                        total_row.append(("", False))
                table_rows.append(total_row)

            story.extend(
                _create_pdf_table(
                    table_rows, headers, normal_style, sec_title, section_style
                )
            )

        # --- Elevation Cost Summary ---
        if elev_sec.get(
            "elevation_cost_summary", elev_sec.get("elevation_summary", True)
        ):
            cs = elev_data.get("cost_summary", {})
            if cs:
                cs_headers = ["Category"]
                cs_col_keys = ["label"]
                if total_count > 1:
                    cs_headers.append("Cost / Elevation")
                    cs_col_keys.append("per_elev")
                cs_headers.append("Total Elevation Cost")
                cs_col_keys.append("total")

                COST_ROWS = [
                    ("Profile Costs", "profile_cost_per_elev", "profile_total_cost"),
                    (
                        "Accessory Costs",
                        "accessory_cost_per_elev",
                        "accessory_total_cost",
                    ),
                    ("Gasket Costs", "gasket_cost_per_elev", "gasket_total_cost"),
                    ("Door Costs", "door_cost_per_elev", "door_total_cost"),
                    ("Glass Costs", "glass_cost_per_elev", "glass_total_cost"),
                    (
                        "Fabrication Costs",
                        "fabrication_cost_per_elev",
                        "fabrication_total_cost",
                    ),
                ]

                cs_table = []
                for label, per_key, total_key in COST_ROWS:
                    total_val = cs.get(total_key, 0)
                    if (
                        total_val == 0
                        and "Door" in label
                        and not elev_data.get("has_doors", False)
                    ):
                        continue
                    row = [label]
                    if total_count > 1:
                        row.append(_fmt_currency(cs.get(per_key, 0)))
                    row.append(_fmt_currency(total_val))
                    cs_table.append(row)

                # Total row
                total_row = [(f"{elev_name} Total", True)]
                if total_count > 1:
                    total_row.append(
                        (_fmt_currency(cs.get("total_cost_per_elev", 0)), True)
                    )
                total_row.append(
                    (_fmt_currency(cs.get("total_elevation_cost", 0)), True)
                )
                cs_table.append(total_row)

                story.extend(
                    _create_pdf_table(
                        cs_table,
                        cs_headers,
                        normal_style,
                        "Elevation Cost Summary",
                        section_style,
                    )
                )

        # --- Diagram ---
        if elev_sec.get("diagram", True):
            diagram_data = elev_data.get("diagram_image")
            if diagram_data and isinstance(diagram_data, bytes):
                try:
                    img_bytes = io.BytesIO(diagram_data)
                    img = Image(img_bytes, width=5 * inch, height=3.75 * inch)
                    img.hAlign = "CENTER"
                    story.append(
                        Paragraph("<b>Bay Distribution Diagram</b>", section_style)
                    )
                    story.append(Spacer(1, 0.1 * inch))
                    story.append(img)
                    story.append(Spacer(1, 0.2 * inch))
                except Exception:
                    pass

        story.append(Spacer(1, 0.3 * inch))

    # =========== SUMMARY SHEET ===========
    if summary_included:
        summary = report_data.get("summary", {})
        if summary:
            story.append(PageBreak())
            story.append(Paragraph("<b>SUMMARY</b>", title_style))
            story.append(Spacer(1, 0.2 * inch))

            sum_tab = summary_opts.get(
                "summary_sections", summary_opts.get("summary_tab", {})
            )
            sum_cols = summary_opts.get("summary_columns", {})
            sum_cost = summary_opts.get("cost_overview", {})

            categories = summary.get("categories", {})
            cat_totals = summary.get("category_totals", {})

            # Category-to-section key mapping
            CAT_SEC_MAP = {
                "PROFILES": "profiles",
                "ACCESSORIES": "accessories",
                "GASKETS": "gaskets",
                "DOORS": "doors",
                "GLASS": "glass",
                "LABOR": "fabrication",
            }

            # --- Material Category Tables ---
            for cat_name, sec_key in CAT_SEC_MAP.items():
                # Check both new key ("fabrication") and legacy ("labor")
                if not sum_tab.get(
                    sec_key,
                    sum_tab.get("labor" if sec_key == "fabrication" else sec_key, True),
                ):
                    continue
                cat_items = categories.get(cat_name, [])
                if not cat_items:
                    continue

                # Build headers based on summary_columns checkboxes + category type
                headers = []
                col_keys = []

                if sum_cols.get("description", True):
                    headers.append("Description")
                    col_keys.append("description")
                if sum_cols.get("project_total_materials", True):
                    headers.append("Project Total Materials")
                    col_keys.append("display")

                # Category-specific quantity columns
                if cat_name == "PROFILES":
                    if sum_cols.get("total_feet_required", True):
                        headers.append("Total Feet Required")
                        col_keys.append("quantity_req_ft")
                    if sum_cols.get("sticks_required", True):
                        headers.append("Sticks Required")
                        col_keys.append("qty_stick_req")
                elif cat_name == "ACCESSORIES":
                    if sum_cols.get("total_pieces_required", True):
                        headers.append("Total Pieces Required")
                        col_keys.append("quantity_req_ft")
                    if sum_cols.get("quantity_per_order", True):
                        headers.append("Qty Per Order")
                        col_keys.append("qty_stick_req")
                    if sum_cols.get("orders_required", True):
                        headers.append("Orders Required")
                        col_keys.append("quantity_display")
                elif cat_name == "GASKETS":
                    if sum_cols.get("total_feet_required", True):
                        headers.append("Total Feet Required")
                        col_keys.append("quantity_req_ft")
                    if sum_cols.get("rolls_required", True):
                        headers.append("Rolls Required")
                        col_keys.append("qty_stick_req")
                else:
                    # GLASS, LABOR, DOORS
                    if sum_cols.get("unit_price", True):
                        headers.append("Unit Price")
                        col_keys.append("qty_stick_req")

                # Common quantity column (profiles, gaskets, glass/labor/doors)
                if cat_name not in ("ACCESSORIES",):
                    if sum_cols.get("total_quantity_required", True):
                        headers.append("Total Quantity Required")
                        col_keys.append("quantity_display")

                if sum_cols.get("total_list_cost", True):
                    headers.append("Total List Cost")
                    col_keys.append("original_total_cost")
                if sum_cols.get("discounted_total_list_cost", True):
                    headers.append("Discounted Total Cost")
                    col_keys.append("total_cost")
                if sum_cols.get("residual_material_quantity", True):
                    headers.append("Residual Qty")
                    col_keys.append("reusable_qty_display")
                if sum_cols.get("residual_waste_pct", True):
                    headers.append("Waste %")
                    col_keys.append("reusable_pct")
                if sum_cols.get("residual_material_cost", True):
                    headers.append("Residual Cost")
                    col_keys.append("reusable_cost")

                if not headers:
                    continue

                table_rows = []
                for item in cat_items:
                    row = []
                    for ck in col_keys:
                        val = item.get(ck, "")
                        if ck in ("original_total_cost", "total_cost", "reusable_cost"):
                            row.append(_fmt_currency(val))
                        elif ck == "reusable_pct":
                            row.append(_fmt_pct(val))
                        else:
                            row.append(str(val) if val else "")
                    table_rows.append(row)

                # Category total row
                ct = cat_totals.get(cat_name, {})
                if ct:
                    total_row = []
                    for ck in col_keys:
                        if ck == "description":
                            total_row.append((f"Total {cat_name.title()} Cost", True))
                        elif ck == "original_total_cost":
                            total_row.append(
                                (_fmt_currency(ct.get("original", 0)), True)
                            )
                        elif ck == "total_cost":
                            total_row.append(
                                (_fmt_currency(ct.get("discounted", 0)), True)
                            )
                        elif ck == "reusable_cost":
                            total_row.append(
                                (_fmt_currency(ct.get("residual", 0)), True)
                            )
                        else:
                            total_row.append(("", False))
                    table_rows.append(total_row)

                story.extend(
                    _create_pdf_table(
                        table_rows, headers, normal_style, cat_name, section_style
                    )
                )

            # --- Elevation Summary Table ---
            if sum_tab.get("elevation_summary", True):
                elev_summ = summary.get("elevation_summary", [])
                if elev_summ:
                    es_headers = [
                        "Elevation",
                        "Quantity",
                        "Dimensions",
                        "SQFT Total",
                        "Perimeter FT",
                    ]
                    es_table = []
                    for es in elev_summ:
                        es_table.append(
                            [
                                es.get("name", ""),
                                str(es.get("quantity", "")),
                                es.get("dimensions", ""),
                                f"{es.get('sqft_total', 0):.2f}",
                                f"{es.get('perimeter_ft_total', 0):.2f}",
                            ]
                        )
                    # Totals row
                    est = summary.get("elevation_summary_totals", {})
                    if est:
                        es_table.append(
                            [
                                ("TOTAL", True),
                                (str(int(est.get("total_qty", 0))), True),
                                ("", False),
                                (f"{est.get('total_sqft', 0):.2f}", True),
                                (f"{est.get('total_perimeter', 0):.2f}", True),
                            ]
                        )
                    story.extend(
                        _create_pdf_table(
                            es_table,
                            es_headers,
                            normal_style,
                            "ELEVATION SUMMARY",
                            section_style,
                        )
                    )

            # --- Cost Overview ---
            cov = summary.get("cost_overview", {})
            if cov:
                co_table = [
                    ["List Price Total", _fmt_currency(cov.get("list_price_total", 0))],
                    [
                        ("Discounted Total", True),
                        (_fmt_currency(cov.get("discounted_total", 0)), True),
                    ],
                    [
                        "Residual / Waste Cost",
                        _fmt_currency(cov.get("residual_waste_cost", 0)),
                    ],
                    ["Waste Percentage", _fmt_pct(cov.get("waste_pct", 0))],
                ]
                story.extend(
                    _create_pdf_table(
                        co_table,
                        ["Item", "Value"],
                        normal_style,
                        "COST OVERVIEW",
                        section_style,
                    )
                )

            # --- Additional Costs ---
            if sum_cost.get("additional_costs", True):
                add_costs = summary.get("additional_costs", {})
                add_items = add_costs.get("items", [])
                add_total = add_costs.get("total", 0)
                if add_items:
                    ac_table = []
                    for label, amount in add_items:
                        ac_table.append([label, _fmt_currency(amount)])
                    ac_table.append(
                        [("SUBTOTAL", True), (_fmt_currency(add_total), True)]
                    )
                    story.extend(
                        _create_pdf_table(
                            ac_table,
                            ["Item", "Amount"],
                            normal_style,
                            "ADDITIONAL COSTS",
                            section_style,
                        )
                    )

            # --- Markups ---
            if sum_cost.get("markups", True):
                markups = summary.get("markups", {})
                mk_items = markups.get("items", [])
                mk_total = markups.get("total", 0)
                if mk_items:
                    mk_table = []
                    for label, amount in mk_items:
                        mk_table.append([label, _fmt_currency(amount)])
                    mk_table.append(
                        [("SUBTOTAL", True), (_fmt_currency(mk_total), True)]
                    )
                    story.extend(
                        _create_pdf_table(
                            mk_table,
                            ["Item", "Amount"],
                            normal_style,
                            "MARKUPS / PROFIT",
                            section_style,
                        )
                    )

            # --- Project Total ---
            pt = summary.get("project_total_breakdown", {})
            if pt:
                pt_table = [
                    ["Discounted Total", _fmt_currency(pt.get("discounted_total", 0))],
                ]
                if pt.get("additional_total", 0) > 0:
                    pt_table.append(
                        ["+ Additional", _fmt_currency(pt.get("additional_total", 0))]
                    )
                if pt.get("markup_total", 0) > 0:
                    pt_table.append(
                        ["+ Markups", _fmt_currency(pt.get("markup_total", 0))]
                    )
                pt_table.append(
                    [
                        ("GRAND TOTAL", True),
                        (_fmt_currency(pt.get("grand_total", 0)), True),
                    ]
                )
                story.extend(
                    _create_pdf_table(
                        pt_table,
                        ["Item", "Amount"],
                        normal_style,
                        "PROJECT TOTAL",
                        section_style,
                    )
                )

            # --- Pie Chart ---
            if sum_cost.get("diagram", True):
                chart_data = summary.get("pie_chart_image")
                if chart_data and isinstance(chart_data, bytes):
                    try:
                        story.append(Spacer(1, 0.4 * inch))
                        story.append(
                            Paragraph("<b>Project Cost Breakdown</b>", section_style)
                        )
                        story.append(Spacer(1, 0.2 * inch))
                        img_bytes = io.BytesIO(chart_data)
                        img = Image(img_bytes, width=5 * inch, height=3.75 * inch)
                        img.hAlign = "CENTER"
                        story.append(img)
                    except Exception:
                        pass

    # Build PDF
    doc.build(story)
    print(f"[OK] PDF exported to: {pdf_path}")
    return pdf_path


def export_project_to_pdf(
    project_name, report_data=None, output_dir="reports", include_logo=True
):
    """Export a project to PDF from report_data dict.

    Args:
        project_name: Project name for filename.
        report_data: The report_data dict. If None, raises error.
        output_dir: Directory for PDF output.
        include_logo: Whether to include company logo.
    Returns:
        Path to generated PDF file.
    """
    if not REPORTLAB_AVAILABLE:
        raise ImportError(
            "reportlab is not installed. Install with: pip install reportlab"
        )

    if report_data is None:
        raise ValueError("report_data is required for PDF generation")

    os.makedirs(output_dir, exist_ok=True)
    pdf_filename = f"{project_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)

    generate_pdf_from_data(report_data, pdf_path, include_logo=include_logo)
    return pdf_path
