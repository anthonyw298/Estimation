# Estimation Script for UNITED GLASS VENTURES

A robust Python-based cost estimation tool to streamline the quoting process for custom glass and door systems. Built with Clear architecture and modular design, this tool combines GUI interaction, part calculations, and Excel report generation.

---

##  Overview

The Estimation Script allows users to:

- Define project configurations using a clean GUI
- Automatically calculate quantities for various parts (glass, gasketing, screws, etc.)
- Easily add, modify, or delete items (e.g., doors, finishes) with consistent pricing logic
- Generate Excel-based estimates with detail-rich breakdowns
- Expand functionality via modular `systems/` and `utils/` directories

---

##  Installation

```bash
git clone https://github.com/anthonyw298/Estimation.git
cd Estimation

python -m venv venv           # optional
venv\Scripts\activate      # Windows

pip install -r requirements.txt
python main.py
