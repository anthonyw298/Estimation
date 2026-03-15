---
paths:
  - "*.py"
  - "utils/**/*.py"
  - "systems/**/*.py"
---

# Python Backend Rules

## main.py
- This is a ~6500 line Flet application -- avoid unnecessary refactors
- Changes should be surgical and targeted
- Test any UI changes by running `python main.py`

## Utils Structure
- `database.py` -- Supabase client singleton (`db`)
- `formulas.py` -- Pure math functions (area, perimeter, doors)
- `excel_generator.py` -- Excel report generation (openpyxl)
- `pdf_generator.py` -- PDF export (reportlab, optional dependency)
- `ml_predictor.py` -- ML predictions (sklearn, optional dependency)
- `waste_calculator.py` -- Waste statistics calculations
- `pricing.py` -- Pricing logic

## Optional Dependencies
ML and PDF features degrade gracefully:
```python
try:
    from utils.pdf_generator import export_project_to_pdf, REPORTLAB_AVAILABLE
except ImportError:
    REPORTLAB_AVAILABLE = False
```
Follow this pattern for any new optional features.

## Data Flow
All project data flows through Supabase. Local `.files/` directory is only for generated Excel reports.
