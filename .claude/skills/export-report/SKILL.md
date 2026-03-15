---
name: export-report
description: Generate or debug Excel/PDF report export functionality
disable-model-invocation: true
allowed-tools: Read, Edit, Write, Grep, Glob, Bash(python *)
---

# Export Report

Work on the report export functionality for `$ARGUMENTS`.

## Context

Report generation code lives in two places:

### Python (Desktop)
- `utils/excel_generator.py` -- Excel reports using openpyxl
- `utils/pdf_generator.py` -- PDF reports using reportlab

### TypeScript (Web)
- `web/src/lib/export.ts` -- Excel export using ExcelJS
- `web/src/lib/pdf-export.ts` -- PDF export using jsPDF + jspdf-autotable
- `web/src/components/ReportOptionsDialog.tsx` -- UI for report options

## Steps

1. Identify which export system is being discussed (Python or TypeScript).

2. Read the relevant files to understand current implementation.

3. Make changes following existing patterns:
   - Match the existing report formatting and styling
   - Handle edge cases (empty data, missing fields)
   - Maintain parity between Python and TypeScript versions where applicable

4. Test the changes if possible.
