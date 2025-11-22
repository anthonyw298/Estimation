import sys
import json
import os
import datetime
from openpyxl import Workbook

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, 
    QTabWidget, QFrame, QScrollArea, QMessageBox, QListWidget, 
    QGroupBox, QGridLayout, QSpacerItem, QSizePolicy
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

# Assuming your utils and systems are in their respective directories
from utils.excel_generator import generate_excel_report
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter, calculate_total_door_area, calculate_glass_to_add_back

# Define a directory to store all project-related files
PROJECTS_DIR = ".files"
MASTER_PROJECT_LIST_FILE = os.path.join(PROJECTS_DIR, "projects_list.json")

class EstimationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("United Glass Estimation Calculation Tool")
        self.resize(1200, 800)

        # Apply dark theme stylesheet
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
                font-family: Helvetica;
            }
            QLineEdit, QComboBox, QListWidget {
                background-color: #3b3b3b;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 4px;
                color: #ffffff;
            }
            QLineEdit:focus, QComboBox:focus {
                border: 1px solid #4a90e2;
            }
            QPushButton {
                background-color: #4a4a4a;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                color: #ffffff;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #5a5a5a;
            }
            QPushButton:pressed {
                background-color: #3a3a3a;
            }
            QPushButton[variant="danger"] {
                background-color: #c62828;
            }
            QPushButton[variant="danger"]:hover {
                background-color: #d32f2f;
            }
            QTabWidget::pane {
                border: 1px solid #444444;
                background-color: #2b2b2b;
            }
            QTabBar::tab {
                background-color: #3b3b3b;
                padding: 8px 16px;
                color: #bbbbbb;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #2b2b2b;
                color: #ffffff;
                border-bottom: 2px solid #4a90e2;
            }
            QGroupBox {
                border: 1px solid #444444;
                border-radius: 6px;
                margin-top: 12px;
                font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
            QLabel[heading="true"] {
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 8px;
            }
            QLabel[subheading="true"] {
                font-size: 14px;
                font-weight: bold;
                margin-top: 8px;
            }
        """)

        # Constants
        self.system_options = ["YES 45TU FRONT SET(OG)", "Other"]
        self.finish_options = ["Clear", "Black", "Paint"]
        self.door_options = ['None', "3' X 7'", "3' X 8'", "3' X 9'", "6' X 7'", "6' X 8'", "6' X 9'"]
        self.stile_options = ["Narrow", "Medium", "Wide"]
        self.hardware_options = [
            "Continuous Hinges", "Concealed Closer", "Exit Devices", "Electric Strike", 
            "Extended Ladder Pull (B2B)", "Extended Ladder Pull (Single)", 
            "Latch Lock w/ Lever Handle", "Lever Handle"
        ]

        # State
        self.saved_elevations = {}
        self.all_projects = []
        self.current_project_name = ""
        self.current_excel_path = ""
        self.current_elevations_json_path = ""
        self.current_extra_materials_json_path = ""
        self.current_door_json_path = ""
        self.selected_door_index = None
        self.doors_data = []

        os.makedirs(PROJECTS_DIR, exist_ok=True)

        # Setup UI
        self.init_ui()
        
        # Initial Load
        self.load_project_list()
        if self.all_projects:
            self.current_project_name = self.all_projects[0]
            self.project_combo.setCurrentText(self.current_project_name)
            self.on_project_select(self.current_project_name)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Tab Widget
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Project Tab
        self.project_tab = QWidget()
        self.setup_project_tab()
        self.tabs.addTab(self.project_tab, "Project Management")

        # Elevation Tab
        self.elevation_tab = QWidget()
        self.setup_elevation_tab()
        self.tabs.addTab(self.elevation_tab, "Elevation Details")

        # Status Bar
        self.status_label = QLabel("")
        self.statusBar().addWidget(self.status_label)

    def setup_project_tab(self):
        layout = QVBoxLayout(self.project_tab)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        title = QLabel("Project Management")
        title.setProperty("heading", True)
        layout.addWidget(title)

        # Create Project Group
        create_group = QGroupBox("Create New Project")
        create_layout = QGridLayout()
        create_group.setLayout(create_layout)
        
        create_layout.addWidget(QLabel("Name:"), 0, 0)
        self.new_project_input = QLineEdit()
        create_layout.addWidget(self.new_project_input, 0, 1)
        
        create_btn = QPushButton("Create")
        create_btn.clicked.connect(self.create_new_project)
        create_layout.addWidget(create_btn, 1, 0, 1, 2)
        
        layout.addWidget(create_group)

        # Manage Project Group
        manage_group = QGroupBox("Manage Existing Projects")
        manage_layout = QGridLayout()
        manage_group.setLayout(manage_layout)
        
        manage_layout.addWidget(QLabel("Select Project:"), 0, 0)
        self.project_combo = QComboBox()
        self.project_combo.currentTextChanged.connect(self.on_project_select)
        manage_layout.addWidget(self.project_combo, 0, 1)
        
        delete_btn = QPushButton("Delete Selected")
        delete_btn.setProperty("variant", "danger")
        delete_btn.clicked.connect(self.delete_current_project)
        manage_layout.addWidget(delete_btn, 1, 0, 1, 2)
        
        layout.addWidget(manage_group)
        layout.addStretch()

    def setup_elevation_tab(self):
        layout = QHBoxLayout(self.elevation_tab)

        # Scroll Area for Left Panel
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # General Details
        general_group = QGroupBox("General Details")
        general_form = QGridLayout()
        general_group.setLayout(general_form)
        
        general_form.addWidget(QLabel("Select System:"), 0, 0)
        self.system_combo = QComboBox()
        self.system_combo.addItems(self.system_options)
        self.system_combo.currentTextChanged.connect(self.on_system_change)
        general_form.addWidget(self.system_combo, 0, 1)
        
        general_form.addWidget(QLabel("Select Finish:"), 1, 0)
        self.finish_combo = QComboBox()
        self.finish_combo.addItems(self.finish_options)
        general_form.addWidget(self.finish_combo, 1, 1)
        
        left_layout.addWidget(general_group)

        # Elevation Specs
        specs_group = QGroupBox("Elevation Specifications")
        specs_form = QGridLayout()
        specs_group.setLayout(specs_form)
        
        specs_form.addWidget(QLabel("Saved Elevations:"), 0, 0)
        self.saved_elevations_combo = QComboBox()
        self.saved_elevations_combo.currentIndexChanged.connect(self.on_saved_elevation_select)
        specs_form.addWidget(self.saved_elevations_combo, 0, 1)
        
        specs_form.addWidget(QLabel("Elevation Type:"), 1, 0)
        self.elevation_type_input = QLineEdit()
        specs_form.addWidget(self.elevation_type_input, 1, 1)
        
        specs_form.addWidget(QLabel("Total Count:"), 2, 0)
        self.total_count_input = QLineEdit()
        specs_form.addWidget(self.total_count_input, 2, 1)
        
        # Bay Inputs (Conditional)
        self.bays_label_w = QLabel("# Bays Wide:")
        self.bays_input_w = QLineEdit()
        self.bays_label_h = QLabel("# Bays Tall:")
        self.bays_input_h = QLineEdit()
        
        self.custom_w_label = QLabel("Custom Bay Widths (csv):")
        self.custom_w_input = QLineEdit()
        self.custom_h_label = QLabel("Custom Bay Heights (csv):")
        self.custom_h_input = QLineEdit()

        specs_form.addWidget(self.bays_label_w, 3, 0)
        specs_form.addWidget(self.bays_input_w, 3, 1)
        specs_form.addWidget(self.bays_label_h, 4, 0)
        specs_form.addWidget(self.bays_input_h, 4, 1)
        
        specs_form.addWidget(self.custom_w_label, 5, 0)
        specs_form.addWidget(self.custom_w_input, 5, 1)
        specs_form.addWidget(self.custom_h_label, 6, 0)
        specs_form.addWidget(self.custom_h_input, 6, 1)

        specs_form.addWidget(QLabel("Opening Width (in):"), 7, 0)
        self.width_input = QLineEdit()
        specs_form.addWidget(self.width_input, 7, 1)
        
        specs_form.addWidget(QLabel("Opening Height (in):"), 8, 0)
        self.height_input = QLineEdit()
        specs_form.addWidget(self.height_input, 8, 1)
        
        left_layout.addWidget(specs_group)

        # Action Buttons
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save Elevation")
        save_btn.clicked.connect(self.save_elevation)
        btn_layout.addWidget(save_btn)
        
        del_elev_btn = QPushButton("Delete Elevation")
        del_elev_btn.setProperty("variant", "danger")
        del_elev_btn.clicked.connect(self.delete_elevation)
        btn_layout.addWidget(del_elev_btn)
        
        gen_report_btn = QPushButton("Generate Report")
        gen_report_btn.clicked.connect(self.generate_report)
        btn_layout.addWidget(gen_report_btn)
        
        left_layout.addLayout(btn_layout)
        left_layout.addStretch()
        
        scroll.setWidget(left_widget)
        layout.addWidget(scroll, 1)

        # Right Panel (Doors)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        door_group = QGroupBox("Door Management")
        door_form = QGridLayout()
        door_group.setLayout(door_form)
        
        door_form.addWidget(QLabel("Door Size:"), 0, 0)
        self.door_size_combo = QComboBox()
        self.door_size_combo.addItems(self.door_options)
        door_form.addWidget(self.door_size_combo, 0, 1)
        
        door_form.addWidget(QLabel("Number of Doors:"), 1, 0)
        self.door_count_input = QLineEdit()
        door_form.addWidget(self.door_count_input, 1, 1)
        
        door_form.addWidget(QLabel("Style:"), 2, 0)
        self.door_style_combo = QComboBox()
        self.door_style_combo.addItems(self.stile_options)
        door_form.addWidget(self.door_style_combo, 2, 1)
        
        door_form.addWidget(QLabel("Hardware:"), 3, 0)
        self.hw_checks = {}
        hw_widget = QWidget()
        hw_layout = QVBoxLayout(hw_widget)
        hw_layout.setContentsMargins(0,0,0,0)
        for opt in self.hardware_options:
            cb = QCheckBox(opt)
            self.hw_checks[opt] = cb
            hw_layout.addWidget(cb)
        door_form.addWidget(hw_widget, 3, 1)

        right_layout.addWidget(door_group)
        
        self.doors_list = QListWidget()
        self.doors_list.itemClicked.connect(self.on_door_select)
        right_layout.addWidget(QLabel("Current Doors:"))
        right_layout.addWidget(self.doors_list)
        
        door_btns = QHBoxLayout()
        add_door_btn = QPushButton("Add Door")
        add_door_btn.clicked.connect(self.add_door)
        door_btns.addWidget(add_door_btn)
        
        update_door_btn = QPushButton("Update Door")
        update_door_btn.clicked.connect(self.update_door)
        door_btns.addWidget(update_door_btn)
        
        del_door_btn = QPushButton("Delete Door")
        del_door_btn.setProperty("variant", "danger")
        del_door_btn.clicked.connect(self.delete_door)
        door_btns.addWidget(del_door_btn)
        
        right_layout.addLayout(door_btns)
        
        layout.addWidget(right_widget, 1)

    # --- Logic ---

    def update_status(self, message, is_error=False):
        color = "#c62828" if is_error else "#ffffff"
        self.status_label.setStyleSheet(f"color: {color}")
        self.status_label.setText(message)

    def load_project_list(self):
        if os.path.exists(MASTER_PROJECT_LIST_FILE):
            try:
                with open(MASTER_PROJECT_LIST_FILE, 'r') as f:
                    self.all_projects = json.load(f)
            except Exception:
                self.all_projects = []
        else:
            self.all_projects = []
        
        self.project_combo.blockSignals(True)
        self.project_combo.clear()
        self.project_combo.addItems(self.all_projects)
        self.project_combo.blockSignals(False)

    def save_project_list(self):
        with open(MASTER_PROJECT_LIST_FILE, 'w') as f:
            json.dump(self.all_projects, f, indent=4)

    def create_new_project(self):
        name = self.new_project_input.text().strip()
        if not name:
            self.update_status("Please enter a project name", True)
            return
        if name in self.all_projects:
            self.update_status("Project already exists", True)
            return
        
        self.all_projects.append(name)
        self.save_project_list()
        self.load_project_list()
        self.project_combo.setCurrentText(name)
        self.new_project_input.clear()
        self.update_status(f"Project '{name}' created")

    def on_project_select(self, text):
        if not text: return
        self.current_project_name = text
        self.set_project_paths()
        self.load_elevations()
        self.clear_form()
        self.update_status(f"Switched to '{text}'")
        self.tabs.setCurrentIndex(1)

    def set_project_paths(self):
        sanitized = self.current_project_name.replace(" ", "_").replace("/", "_").replace("\\", "_")
        base = os.path.join(PROJECTS_DIR, sanitized)
        self.current_excel_path = f"{base}_Report.xlsx"
        self.current_elevations_json_path = f"{base}_Elevations.json"
        self.current_extra_materials_json_path = f"{base}_ExtraMaterials.json"
        
        if not os.path.exists(self.current_excel_path):
            wb = Workbook()
            ws = wb.active
            ws.title = "Report"
            wb.save(self.current_excel_path)
            
        if not os.path.exists(self.current_extra_materials_json_path):
            with open(self.current_extra_materials_json_path, 'w') as f:
                json.dump({}, f)

    def load_elevations(self):
        if os.path.exists(self.current_elevations_json_path):
            try:
                with open(self.current_elevations_json_path, 'r') as f:
                    self.saved_elevations = json.load(f)
            except:
                self.saved_elevations = {}
        else:
            self.saved_elevations = {}
            with open(self.current_elevations_json_path, 'w') as f:
                json.dump({}, f)
        
        self.saved_elevations_combo.blockSignals(True)
        self.saved_elevations_combo.clear()
        elevs = sorted(self.saved_elevations.keys())
        self.saved_elevations_combo.addItems(elevs)
        self.saved_elevations_combo.blockSignals(False)
        
        if elevs:
            self.saved_elevations_combo.setCurrentText(elevs[0])
            self.on_saved_elevation_select(0)

    def on_saved_elevation_select(self, index):
        elev_type = self.saved_elevations_combo.currentText()
        if not elev_type or elev_type not in self.saved_elevations:
            self.clear_form()
            return

        data = self.saved_elevations[elev_type]
        
        self.system_combo.setCurrentText(data.get('system', self.system_options[0]))
        self.finish_combo.setCurrentText(data.get('finish', self.finish_options[0]))
        self.elevation_type_input.setText(elev_type)
        self.total_count_input.setText(str(data.get('total_count', '')))
        self.width_input.setText(str(data.get('opening_width_inches', '')))
        self.height_input.setText(str(data.get('opening_height_inches', '')))
        
        self.bays_input_w.setText(str(data.get('bays_wide', '')))
        self.bays_input_h.setText(str(data.get('bays_tall', '')))
        
        cw = data.get('custom_bay_widths', [])
        self.custom_w_input.setText(','.join(map(str, cw)) if cw else "")
        
        ch = data.get('custom_bay_heights', [])
        self.custom_h_input.setText(','.join(map(str, ch)) if ch else "")

        self.current_door_json_path = self._ensure_door_file(elev_type)
        self.load_doors()
        self.on_system_change(self.system_combo.currentText())

    def on_system_change(self, text):
        visible = text == "YES 45TU FRONT SET(OG)"
        self.bays_label_w.setVisible(visible)
        self.bays_input_w.setVisible(visible)
        self.bays_label_h.setVisible(visible)
        self.bays_input_h.setVisible(visible)
        self.custom_w_label.setVisible(visible)
        self.custom_w_input.setVisible(visible)
        self.custom_h_label.setVisible(visible)
        self.custom_h_input.setVisible(visible)

    def _ensure_door_file(self, elev_type):
        if not self.current_project_name or not elev_type: return None
        proj = self.current_project_name.replace(" ", "_")
        safe_elev = elev_type.replace(" ", "_")
        path = os.path.join(PROJECTS_DIR, f"{proj}_{safe_elev}_doors.json")
        
        if not os.path.exists(path):
            with open(path, 'w') as f: json.dump([], f)
        return path

    def load_doors(self):
        self.doors_list.clear()
        self.doors_data = []
        if not self.current_door_json_path or not os.path.exists(self.current_door_json_path):
            return
        
        try:
            with open(self.current_door_json_path, 'r') as f:
                self.doors_data = json.load(f)
        except: return

        for i, door in enumerate(self.doors_data):
            hw = [k for k,v in door['hardware'].items() if v]
            hw_str = f" - HW: {', '.join(hw)}" if hw else ""
            self.doors_list.addItem(f"Door {i+1}: {door['size']}, {door['stile']}, Qty: {door['count']}{hw_str}")

    def on_door_select(self, item):
        idx = self.doors_list.row(item)
        self.selected_door_index = idx
        if 0 <= idx < len(self.doors_data):
            d = self.doors_data[idx]
            self.door_size_combo.setCurrentText(d['size'])
            self.door_count_input.setText(str(d['count']))
            self.door_style_combo.setCurrentText(d['stile'])
            for k, cb in self.hw_checks.items():
                cb.setChecked(d['hardware'].get(k, False))

    def clear_form(self):
        self.elevation_type_input.clear()
        self.total_count_input.clear()
        self.width_input.clear()
        self.height_input.clear()
        self.bays_input_w.clear()
        self.bays_input_h.clear()
        self.custom_w_input.clear()
        self.custom_h_input.clear()
        self.clear_door_form()
        self.doors_list.clear()
        self.doors_data = []

    def clear_door_form(self):
        self.door_size_combo.setCurrentIndex(0)
        self.door_count_input.clear()
        self.door_style_combo.setCurrentIndex(0)
        for cb in self.hw_checks.values(): cb.setChecked(False)
        self.selected_door_index = None
        self.doors_list.clearSelection()

    def parse_custom_bays(self, text, total, count):
        if not text.strip(): 
            return [total/count]*count if count else []
        try:
            vals = [float(x) for x in text.split(',') if x.strip()]
            if len(vals) > count: raise ValueError
            if sum(vals) > total: raise ValueError
            rem = count - len(vals)
            if rem > 0:
                rem_val = (total - sum(vals)) / rem
                vals += [rem_val] * rem
            return vals
        except:
            return [total/count]*count if count else []

    def save_elevation(self):
        if not self.current_project_name:
            self.update_status("No project selected", True)
            return

        try:
            elev = self.elevation_type_input.text().strip()
            if not elev: raise ValueError("Missing elevation type")
            
            total = int(self.total_count_input.text())
            w = float(self.width_input.text())
            h = float(self.height_input.text())
            
            sqft = calculate_rectangle_area(w/12, h/12)
            perim = calculate_perimeter(w/12, h/12)
            
            self.current_door_json_path = self._ensure_door_file(elev)
            
            data = {
                'system': self.system_combo.currentText(),
                'finish': self.finish_combo.currentText(),
                'total_count': total,
                'opening_width_inches': w,
                'opening_height_inches': h,
                'sqft_per_type': sqft,
                'total_sqft': sqft * total,
                'perimeter_ft': perim,
                'total_perimeter_ft': perim * total
            }
            
            outputs = []
            if data['system'] == "YES 45TU FRONT SET(OG)":
                bw = int(self.bays_input_w.text())
                bh = int(self.bays_input_h.text())
                cw = self.parse_custom_bays(self.custom_w_input.text(), w, bw)
                ch = self.parse_custom_bays(self.custom_h_input.text(), h, bh)
                
                data.update({'bays_wide': bw, 'bays_tall': bh, 
                           'custom_bay_widths': cw, 'custom_bay_heights': ch})
                
                outputs = calculate_yes45tu_quantities(bw, bh, total, w, h, self.doors_data)

            data['calculated_outputs'] = outputs
            self.saved_elevations[elev] = data
            
            with open(self.current_elevations_json_path, 'w') as f:
                json.dump(self.saved_elevations, f, indent=4)
            
            generate_excel_report(
                self.current_excel_path, self.current_elevations_json_path,
                self.current_extra_materials_json_path, data['system'], data['finish'],
                elev, total, data.get('bays_wide', 0), data.get('bays_tall', 0),
                w, h, sqft, sqft*total, perim, perim*total, outputs, None, 
                self.doors_data, data.get('custom_bay_widths', []), 
                data.get('custom_bay_heights', [])
            )
            
            self.load_elevations()
            self.saved_elevations_combo.setCurrentText(elev)
            self.update_status(f"Saved elevation '{elev}'")
            
        except ValueError as e:
            self.update_status(f"Error: {e}", True)
        except Exception as e:
            self.update_status(f"Unexpected error: {e}", True)

    def delete_elevation(self):
        elev = self.saved_elevations_combo.currentText()
        if elev in self.saved_elevations:
            del self.saved_elevations[elev]
            with open(self.current_elevations_json_path, 'w') as f:
                json.dump(self.saved_elevations, f, indent=4)
            
            dp = self._ensure_door_file(elev)
            if dp and os.path.exists(dp): os.remove(dp)
            
            self.load_elevations()
            self.clear_form()
            self.update_status("Elevation deleted")

    def generate_report(self):
        if not self.current_project_name: return
        
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        pname = self.current_project_name.replace(" ", "_")
        path = os.path.join("reports", f"{pname}_Report_{ts}.xlsx")
        os.makedirs("reports", exist_ok=True)
        
        try:
            generate_excel_report(
                path, self.current_elevations_json_path, 
                self.current_extra_materials_json_path, "", "", "", 0,0,0,0,0,0,0,0,0,
                [], None, None, mode="export_all"
            )
            self.update_status(f"Report generated: {path}")
        except Exception as e:
            self.update_status(f"Report error: {e}", True)

    def delete_current_project(self):
        if not self.current_project_name: return
        
        reply = QMessageBox.question(self, 'Confirm Delete', 
            f"Delete project '{self.current_project_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            
        if reply == QMessageBox.StandardButton.Yes:
            # Simple deletion logic
            pname = self.current_project_name.replace(" ", "_")
            files = [
                f"{pname}_Report.xlsx", f"{pname}_Elevations.json", 
                f"{pname}_ExtraMaterials.json"
            ]
            for f in os.listdir(PROJECTS_DIR):
                if f.startswith(pname) and "_doors.json" in f:
                    files.append(f)
            
            for f in files:
                fp = os.path.join(PROJECTS_DIR, f)
                if os.path.exists(fp): os.remove(fp)
                
            self.all_projects.remove(self.current_project_name)
            self.save_project_list()
            self.load_project_list()
            self.clear_form()
            self.update_status("Project deleted")

    def add_door(self): self._modify_door('add')
    def update_door(self): self._modify_door('update')
    
    def delete_door(self):
        if self.selected_door_index is None: return
        del self.doors_data[self.selected_door_index]
        self._save_doors()
        self.update_status("Door deleted")

    def _modify_door(self, action):
        try:
            count = int(self.door_count_input.text())
            if count <= 0: raise ValueError
        except:
            self.update_status("Invalid door count", True)
            return

        if not self.elevation_type_input.text():
            self.update_status("Enter elevation type first", True)
            return

        new_door = {
            'size': self.door_size_combo.currentText(),
            'count': count,
            'stile': self.door_style_combo.currentText(),
            'hardware': {k: cb.isChecked() for k, cb in self.hw_checks.items()}
        }

        doors = self.doors_data.copy()
        if action == 'add': doors.append(new_door)
        elif self.selected_door_index is not None:
            doors[self.selected_door_index] = new_door
        else: return

        # Simple validation check logic here...
        
        self.doors_data = doors
        self._save_doors()
        self.save_elevation()
        self.clear_door_form()

    def _save_doors(self):
        if not self.current_door_json_path:
            self.current_door_json_path = self._ensure_door_file(self.elevation_type_input.text())
        
        with open(self.current_door_json_path, 'w') as f:
            json.dump(self.doors_data, f, indent=4)
        self.load_doors()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EstimationApp()
    window.show()
    sys.exit(app.exec())

