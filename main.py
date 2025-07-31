import customtkinter as ctk
import tkinter as tk
import json
import os
from openpyxl import Workbook
import tkinter.messagebox

# Assuming your utils are in a 'utils' directory relative to this script
from utils.excel_generator import generate_excel_report, create_summary_sheet
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter

# Define a directory to store all project-related files
PROJECTS_DIR = "projects"
MASTER_PROJECT_LIST_FILE = os.path.join(PROJECTS_DIR, "projects_list.json")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("United Glass Estimation Calculation Tool")
        self.state('zoomed')
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        os.makedirs(PROJECTS_DIR, exist_ok=True)

        self.system_options = ["YES 45TU FRONT SET(OG)", "Other"]
        self.finish_options = ["Clear", "Black", "Paint"]
        self.door_options = ['None', "3' X 7'", "3' X 8'", "3' X 9'", "6' X 7'", "6' X 8'", "6' X 9'"]
        self.stile_options = ["Narrow", "Medium", "Wide"]
        self.hardware_options = [
            "Continuous Hinges", "Concealed Closer", "Exit Devices", "Electric Strike", 
            "Extended Ladder Pull (B2B)", "Extended Ladder Pull (Single)", 
            "Latch Lock w/ Lever Handle", "Lever Handle"
        ]
        self.saved_elevations = {}
        self.all_projects = []
        self.current_project_name = ""
        self.current_excel_path = ""
        self.current_elevations_json_path = ""
        self.current_extra_materials_json_path = ""

        self.vars = dict(
            system=tk.StringVar(value=self.system_options[0]),
            finish=tk.StringVar(value=self.finish_options[0]),
            door=tk.StringVar(value=self.door_options[0]),
            door_count=tk.StringVar(value=""),
            stile=tk.StringVar(value=self.stile_options[0]),
            
            elevation_type=tk.StringVar(),
            total_count=tk.StringVar(),
            bays_wide=tk.StringVar(),
            bays_tall=tk.StringVar(),
            opening_width=tk.StringVar(),
            opening_height=tk.StringVar(),
            saved_elevation_types=tk.StringVar(),
            new_project_name=tk.StringVar(),
            selected_project=tk.StringVar()
        )
        self.hardware_vars = {opt: tk.BooleanVar(value=False) for opt in self.hardware_options}
        self.widgets = {}

        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.tab_view = ctk.CTkTabview(self.main_frame)
        self.tab_view.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.project_tab = self.tab_view.add("Project Management")
        self.elevation_tab = self.tab_view.add("Elevation Details")
        
        self.create_project_tab_widgets()
        self.create_elevation_tab_widgets()
        
        self.load_project_list()
        if self.all_projects:
            self.vars['selected_project'].set(self.all_projects[0])
            self.on_project_select(self.all_projects[0])
        else:
            self.update_project_dropdown()
            self.update_status("Info: No projects found. Create a new one.", "blue")
            
        self.on_system_change(self.vars['system'].get())
        self.on_door_change(self.vars['door'].get())

    def create_project_tab_widgets(self):
        """Builds the UI for the Project Management tab."""
        self.project_tab.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.project_tab, text="Project Management", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, columnspan=3, pady=(10, 20), sticky="ew")

        # New Project Section
        ctk.CTkLabel(self.project_tab, text="New Project Name:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        new_project_entry = ctk.CTkEntry(self.project_tab, textvariable=self.vars['new_project_name'])
        new_project_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        ctk.CTkButton(self.project_tab, text="Create", command=self.create_new_project).grid(row=1, column=2, padx=10, pady=5)
        
        # Existing Project Section
        ctk.CTkLabel(self.project_tab, text="Select Project:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.project_dropdown = ctk.CTkOptionMenu(self.project_tab, values=[], variable=self.vars['selected_project'], command=self.on_project_select)
        self.project_dropdown.grid(row=2, column=1, sticky="ew", padx=10, pady=5)
        ctk.CTkButton(self.project_tab, text="Delete Selected Project", fg_color="red", hover_color="darkred", command=self.delete_current_project).grid(row=2, column=2, padx=10, pady=5)
        
        self.status_label = ctk.CTkLabel(self.project_tab, text="", text_color="orange")
        self.status_label.grid(row=3, column=0, columnspan=3, pady=10)

    def create_elevation_tab_widgets(self):
        """Builds the UI for the Elevation Details tab."""
        self.elevation_tab.grid_columnconfigure(0, weight=1)
        self.elevation_tab.grid_rowconfigure(0, weight=1) 

        self.scroll_frame = ctk.CTkScrollableFrame(self.elevation_tab)
        self.scroll_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.scroll_frame.grid_columnconfigure(1, weight=1)
        
        self.scroll_row = 0

        # Group 1: General Details
        self._add_header(self.scroll_frame, "General Details", name="general_details_header")
        self._add_input_row(self.scroll_frame, "Select System:", ctk.CTkOptionMenu, self.vars['system'], self.system_options, command=self.on_system_change)
        self._add_input_row(self.scroll_frame, "Select Finish:", ctk.CTkOptionMenu, self.vars['finish'], self.finish_options)
        
        # Group 2: Door Details
        self._add_header(self.scroll_frame, "Door Details", name="door_details_header")
        self._add_input_row(self.scroll_frame, "Select Door Size:", ctk.CTkOptionMenu, self.vars['door'], self.door_options, command=self.on_door_change)
        
        self.door_details_frame = ctk.CTkFrame(self.scroll_frame, fg_color="gray20", corner_radius=10)
        self.door_details_frame.grid_columnconfigure(1, weight=1)
        self._create_door_widgets(self.door_details_frame)

        # Group 3: Elevation Specifications
        self._add_header(self.scroll_frame, "Elevation Specifications", name="elevation_spec_header")
        
        self.saved_elevations_label, self.saved_elevations_option_menu = self._add_input_row(self.scroll_frame, "Saved Elevations:", ctk.CTkOptionMenu, self.vars['saved_elevation_types'], [], command=self.on_saved_elevation_select)
        self.elevation_type_label, self.elevation_type_entry = self._add_input_row(self.scroll_frame, "Elevation Type:", ctk.CTkEntry, self.vars['elevation_type'], None)
        self.total_count_label, self.total_count_entry = self._add_input_row(self.scroll_frame, "Total Count:", ctk.CTkEntry, self.vars['total_count'], None)
        self.bays_wide_label, self.bays_wide_entry = self._add_input_row(self.scroll_frame, "# Bays Wide:", ctk.CTkEntry, self.vars['bays_wide'], None)
        self.bays_tall_label, self.bays_tall_entry = self._add_input_row(self.scroll_frame, "# Bays Tall:", ctk.CTkEntry, self.vars['bays_tall'], None)
        self.opening_width_label, self.opening_width_entry = self._add_input_row(self.scroll_frame, "Opening Width (in):", ctk.CTkEntry, self.vars['opening_width'], None)
        self.opening_height_label, self.opening_height_entry = self._add_input_row(self.scroll_frame, "Opening Height (in):", ctk.CTkEntry, self.vars['opening_height'], None)
        
        self.button_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.button_frame.grid(row=self.scroll_row, column=0, columnspan=2, pady=20)
        ctk.CTkButton(self.button_frame, text="Save Elevation", command=self.save_elevation_type).pack(side="left", padx=10)
        ctk.CTkButton(self.button_frame, text="Delete Elevation", fg_color="red", hover_color="darkred", command=self.delete_elevation_type).pack(side="left", padx=10)
        self.scroll_row += 1
        
        # After creating all widgets, re-grid them to their correct initial positions
        self._reposition_widgets()


    def _add_header(self, parent, text, name=None):
        """Helper to add a bold section header."""
        lbl = ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=16, weight="bold"))
        self.widgets[name] = lbl
        self.scroll_row += 1
        return lbl

    def _add_input_row(self, parent, label_text, widget_type, variable, options, **kwargs):
        """Helper to create a label and input widget in a grid row, and increment the row counter."""
        lbl = ctk.CTkLabel(parent, text=label_text)
        
        if widget_type == ctk.CTkEntry:
            widget = widget_type(parent, textvariable=variable, **kwargs)
        elif widget_type == ctk.CTkOptionMenu:
            widget = widget_type(parent, values=options if options else [""], variable=variable, **kwargs)
        
        widget_id = label_text.replace(" ", "_").replace(":", "").lower()
        self.widgets[f"label_{widget_id}"] = lbl
        self.widgets[f"widget_{widget_id}"] = widget

        self.scroll_row += 1
        return lbl, widget

    def _create_door_widgets(self, parent_frame):
        """Creates door-specific widgets inside a frame."""
        self.door_details_row = 0
        self._add_input_row_to_frame(parent_frame, "Number of Doors:", ctk.CTkEntry, self.vars['door_count'])
        self._add_input_row_to_frame(parent_frame, "Stile Style:", ctk.CTkOptionMenu, self.vars['stile'], self.stile_options, command=self.on_stile_change)
        
        hardware_label = ctk.CTkLabel(parent_frame, text="Select Hardware:")
        hardware_label.grid(row=self.door_details_row, column=0, sticky="w", padx=10, pady=5)
        self.door_details_row += 1

        self.hardware_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        self.hardware_frame.grid(row=self.door_details_row, column=0, columnspan=2, sticky="ew", padx=10)
        self.door_details_row += 1
        
        row_idx = 0
        for option in self.hardware_options:
            cb = ctk.CTkCheckBox(self.hardware_frame, text=option, variable=self.hardware_vars[option])
            cb.grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
            row_idx += 1
            
        self.door_details_frame.grid_forget()

    def _add_input_row_to_frame(self, parent, label_text, widget_type, variable, options=None, **kwargs):
        """Helper to add a label and input widget to a specified frame."""
        lbl = ctk.CTkLabel(parent, text=label_text)
        lbl.grid(row=self.door_details_row, column=0, sticky="w", padx=10, pady=5)
        
        if widget_type == ctk.CTkEntry:
            widget = widget_type(parent, textvariable=variable, **kwargs)
        elif widget_type == ctk.CTkOptionMenu:
            widget = widget_type(parent, values=options if options else [""], variable=variable, **kwargs)
        
        widget.grid(row=self.door_details_row, column=1, sticky="ew", padx=10, pady=5)
        self.door_details_row += 1
    
    def _reposition_widgets(self):
        """Helper to dynamically re-grid all widgets in the correct order."""
        current_row = 0
        
        # General Details
        self.widgets['general_details_header'].grid(row=current_row, column=0, columnspan=2, sticky="w", pady=(20, 5), padx=5); current_row += 1
        self.widgets['label_select_system'].grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        self.widgets['widget_select_system'].grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
        self.widgets['label_select_finish'].grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        self.widgets['widget_select_finish'].grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
        
        # Door Details
        self.widgets['door_details_header'].grid(row=current_row, column=0, columnspan=2, sticky="w", pady=(20, 5), padx=5); current_row += 1
        self.widgets['label_select_door_size'].grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        self.widgets['widget_select_door_size'].grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
        
        # Place door details frame if a door is selected
        if self.vars['door'].get() != 'None':
            self.door_details_frame.grid(row=current_row, column=0, columnspan=2, sticky="ew", padx=10, pady=10)
            current_row += 1
        
        # Elevation Specifications
        self.widgets['elevation_spec_header'].grid(row=current_row, column=0, columnspan=2, sticky="w", pady=(20, 5), padx=5); current_row += 1
        
        self.saved_elevations_label.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        self.saved_elevations_option_menu.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
        self.elevation_type_label.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        self.elevation_type_entry.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
        self.total_count_label.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        self.total_count_entry.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1

        show_bays = self.vars['system'].get() == "YES 45TU FRONT SET(OG)"
        if show_bays:
            self.bays_wide_label.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
            self.bays_wide_entry.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
            self.bays_tall_label.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
            self.bays_tall_entry.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
        else:
            self.bays_wide_label.grid_forget()
            self.bays_wide_entry.grid_forget()
            self.bays_tall_label.grid_forget()
            self.bays_tall_entry.grid_forget()
            
        self.opening_width_label.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        self.opening_width_entry.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
        self.opening_height_label.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        self.opening_height_entry.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5); current_row += 1
        
        # Buttons
        self.button_frame.grid(row=current_row, column=0, columnspan=2, pady=20)

    # --- UI Logic Methods ---
    def on_door_change(self, selected_door):
        is_door_selected = selected_door != "None"
        if is_door_selected:
            self.on_stile_change(self.vars['stile'].get())
        else:
            self.door_details_frame.grid_forget()
            self.vars['door_count'].set("")
            self.vars['stile'].set(self.stile_options[0])
            for var in self.hardware_vars.values():
                var.set(False)
        self._reposition_widgets()

    def on_stile_change(self, selected_stile):
        is_stile_selected = selected_stile in self.stile_options
        if is_stile_selected:
            self.hardware_frame.grid()
        else:
            self.hardware_frame.grid_forget()
        # Repositioning is handled by on_door_change, which is the main trigger

    def on_system_change(self, selected):
        # Repositioning is handled by the main reposition function
        self._reposition_widgets()
    
    def on_saved_elevation_select(self, elev_type):
        if elev_type not in self.saved_elevations:
            return
        data = self.saved_elevations[elev_type]
        
        for var in self.hardware_vars.values():
            var.set(False)

        for key, var_key in [
            ('system', 'system'), ('finish', 'finish'), ('door_size', 'door'),
            ('door_count', 'door_count'), ('door_stile_style', 'stile'),
            ('total_count', 'total_count'), ('bays_wide', 'bays_wide'),
            ('bays_tall', 'bays_tall'), ('opening_width_inches', 'opening_width'),
            ('opening_height_inches', 'opening_height'),
        ]:
            self.vars[var_key].set(str(data.get(key, '')))
        
        for hardware_item in data.get('door_hardware', []):
            if hardware_item in self.hardware_vars:
                self.hardware_vars[hardware_item].set(True)

        self.vars['elevation_type'].set(elev_type)
        self.on_system_change(self.vars['system'].get())
        self.on_door_change(self.vars['door'].get())
        self.on_stile_change(self.vars['stile'].get())
        self.update_status(f"Loaded elevation: {elev_type}", "green")
        self._reposition_widgets()

    def clear_form(self):
        for var in ['elevation_type', 'total_count', 'bays_wide', 'bays_tall', 'opening_width', 'opening_height', 'door_count']:
            self.vars[var].set("")
        self.vars['system'].set(self.system_options[0])
        self.vars['finish'].set(self.finish_options[0])
        self.vars['door'].set(self.door_options[0])
        self.vars['stile'].set(self.stile_options[0])
        for var in self.hardware_vars.values():
            var.set(False)
        self.on_system_change(self.vars['system'].get())
        self.on_door_change(self.vars['door'].get())
        
    def update_status(self, message, color="orange"):
        self.status_label.configure(text=message, text_color=color)

    # --- Project Management Methods ---
    def load_project_list(self):
        if os.path.exists(MASTER_PROJECT_LIST_FILE):
            try:
                with open(MASTER_PROJECT_LIST_FILE, 'r') as f:
                    self.all_projects = json.load(f)
            except Exception:
                self.all_projects = []
        else:
            self.all_projects = []
        self.update_project_dropdown()

    def save_project_list(self):
        with open(MASTER_PROJECT_LIST_FILE, 'w') as f:
            json.dump(self.all_projects, f, indent=4)

    def update_project_dropdown(self):
        values = self.all_projects if self.all_projects else [""]
        self.project_dropdown.configure(values=values)
        if self.all_projects:
            self.vars['selected_project'].set(self.current_project_name if self.current_project_name in self.all_projects else self.all_projects[0])
        else:
            self.vars['selected_project'].set("")

    def create_new_project(self):
        new_name = self.vars['new_project_name'].get().strip()
        if not new_name:
            self.update_status("Error: Please enter a name for the new project.", "red")
            return
        if new_name in self.all_projects:
            self.update_status(f"Error: Project '{new_name}' already exists.", "red")
            return
        
        self.all_projects.append(new_name)
        self.save_project_list()
        self.update_project_dropdown()
        self.vars['selected_project'].set(new_name)
        self.on_project_select(new_name)
        self.update_status(f"Project '{new_name}' created.", "green")
        self.vars['new_project_name'].set("")
        self.tab_view.set("Elevation Details")

    def on_project_select(self, project_name):
        self.current_project_name = project_name
        self.set_current_project_paths()
        self.load_saved_elevations_for_current_project()
        self.clear_form()
        self.update_status(f"Switched to project '{project_name}'", "blue")

    def set_current_project_paths(self):
        if not self.current_project_name: return
        base_name = os.path.join(PROJECTS_DIR, self.current_project_name.replace(" ", "_").replace("/", "_").replace("\\", "_"))
        self.current_excel_path = f"{base_name}_Report.xlsx"
        self.current_elevations_json_path = f"{base_name}_Elevations.json"
        self.current_extra_materials_json_path = f"{base_name}_ExtraMaterials.json"
        
        if not os.path.exists(self.current_excel_path):
            Workbook().save(self.current_excel_path)
        if not os.path.exists(self.current_extra_materials_json_path):
            with open(self.current_extra_materials_json_path, 'w') as f:
                json.dump({}, f, indent=4)
    
    def load_saved_elevations_for_current_project(self):
        if os.path.exists(self.current_elevations_json_path):
            try:
                with open(self.current_elevations_json_path, 'r') as f:
                    self.saved_elevations = json.load(f)
            except Exception:
                self.saved_elevations = {}
        else:
            self.saved_elevations = {}
            with open(self.current_elevations_json_path, 'w') as f:
                json.dump({}, f, indent=4)
        self.update_saved_elevation_dropdown()
    
    def update_saved_elevation_dropdown(self):
        keys = sorted(self.saved_elevations.keys())
        self.widgets['widget_saved_elevations'].configure(values=keys if keys else [""])
        current = self.vars['saved_elevation_types'].get()
        if current not in keys:
            self.vars['saved_elevation_types'].set(keys[0] if keys else "")

    def delete_current_project(self):
        if not self.current_project_name:
            self.update_status("Error: No project selected to delete.", "red")
            return
        
        if not tk.messagebox.askyesno(
            "Confirm Deletion", f"Are you sure you want to delete project '{self.current_project_name}' and all its data?"
        ):
            self.update_status("Project deletion cancelled.", "blue")
            return
            
        try:
            files_to_delete = [self.current_excel_path, self.current_elevations_json_path, self.current_extra_materials_json_path]
            for file_path in files_to_delete:
                if os.path.exists(file_path): os.remove(file_path)
            
            if self.current_project_name in self.all_projects:
                self.all_projects.remove(self.current_project_name)
                self.save_project_list()
                self.update_project_dropdown()
            
            self.update_status(f"Project '{self.current_project_name}' and its files deleted.", "green")
            if self.all_projects:
                self.on_project_select(self.all_projects[0])
            else:
                self.vars['selected_project'].set("")
                self.current_project_name = ""
                self.clear_form()
                self.saved_elevations = {}
                self.update_saved_elevation_dropdown()
                self.update_status("Info: No projects remaining. Create a new one.", "blue")
        except Exception as e:
            self.update_status(f"Error: Could not delete project '{self.current_project_name}': {e}", "red")

    def save_elevation_type(self):
        if not self.current_project_name:
            self.update_status("Error: Please create or select a project first.", "red")
            return
        try:
            v = self.vars
            elev = v['elevation_type'].get().strip()
            if not elev:
                self.update_status("Error: Please enter an elevation type.", "red")
                return
            system = v['system'].get()
            finish = v['finish'].get()
            door_size = v['door'].get()
            door_count = int(v['door_count'].get()) if door_size != 'None' and v['door_count'].get() else 0
            stile_style = v['stile'].get() if door_count > 0 else 'None'
            door_hardware = [opt for opt, var in self.hardware_vars.items() if var.get()]

            total = int(v['total_count'].get())
            ow = float(v['opening_width'].get())
            oh = float(v['opening_height'].get())
            bays_wide = int(v['bays_wide'].get()) if system == self.system_options[0] and v['bays_wide'].get() else 0
            bays_tall = int(v['bays_tall'].get()) if system == self.system_options[0] and v['bays_tall'].get() else 0

            calculated = []
            if system == self.system_options[0]:
                calculated = calculate_yes45tu_quantities(
                    bays_wide, bays_tall, total, ow, oh, door_size, door_count, stile_style, door_hardware
                )

            sqft_per = calculate_rectangle_area(ow / 12, oh / 12)
            total_sqft = sqft_per * total
            perimeter = calculate_perimeter(ow / 12, oh / 12)
            total_perimeter = perimeter * total

            generate_excel_report(
                excel_path=self.current_excel_path,
                elevations_json_path=self.current_elevations_json_path,
                extra_materials_json_path=self.current_extra_materials_json_path,
                system_input=system, finish_input=finish, elevation_type=elev,
                total_count=total, bays_wide=bays_wide, bays_tall=bays_tall,
                opening_width=ow, opening_height=oh, sqft_per_type=sqft_per,
                total_sqft=total_sqft, perimeter_ft=perimeter, total_perimeter_ft=total_perimeter,
                calculated_outputs=calculated, door_size=door_size, door_count=door_count,
                stile_style=stile_style, door_hardware=door_hardware, mode="save_or_update"
            )
            
            self.load_saved_elevations_for_current_project()
            self.vars['saved_elevation_types'].set(elev)
            self.update_status(f"Elevation '{elev}' saved successfully.", "green")

        except ValueError as e:
            self.update_status(f"Error: {e}", "red")
        except Exception as e:
            self.update_status(f"An unexpected error occurred: {e}", "red")

    def delete_elevation_type(self):
        if not self.current_project_name:
            self.update_status("Error: Please select a project first.", "red")
            return
        elev = self.vars['saved_elevation_types'].get()
        if not elev:
            self.update_status("Error: No elevation selected to delete.", "red")
            return
            
        if not tk.messagebox.askyesno(
            "Confirm Deletion", f"Are you sure you want to delete elevation '{elev}'?"
        ):
            self.update_status("Deletion cancelled.", "blue")
            return
            
        generate_excel_report(
            excel_path=self.current_excel_path,
            elevations_json_path=self.current_elevations_json_path,
            extra_materials_json_path=self.current_extra_materials_json_path,
            delete_elevation_type=elev
        )
        self.load_saved_elevations_for_current_project()
        self.clear_form()
        self.update_status(f"Elevation '{elev}' deleted successfully.", "green")

if __name__ == "__main__":
    app = App()
    app.after(10, lambda: app.state('zoomed'))
    app.mainloop()