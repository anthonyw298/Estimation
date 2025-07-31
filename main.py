import customtkinter as ctk
import tkinter as tk
import json
import os
from openpyxl import Workbook
import tkinter.messagebox
import re

# Define a directory to store all project-related files
PROJECTS_DIR = "projects"
MASTER_PROJECT_LIST_FILE = os.path.join(PROJECTS_DIR, "projects_list.json")

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("United Glass Estimation Calculation Tool")
        self.state('zoomed')
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")
        
        self.bg_color = "#2b2b2b"
        self.fg_color = "#1f1f1f"
        self.text_color = "#ffffff"
        self.accent_color = "#555555"
        self.accent_hover = "#666666"
        self.border_color = "#444444"
        self.success_color = "#2e7d32"
        self.error_color = "#c62828"
        
        self.configure(fg_color=self.bg_color)
        
        self.font_family = "Helvetica"
        self.title_font = ctk.CTkFont(family=self.font_family, size=24, weight="bold")
        self.header_font = ctk.CTkFont(family=self.font_family, size=18, weight="bold")
        self.label_font = ctk.CTkFont(family=self.font_family, size=14)
        self.button_font = ctk.CTkFont(family=self.font_family, size=14, weight="bold")

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
        
        self.current_elevation_doors = []
        self.selected_door_index = None

        self.vars = dict(
            system=tk.StringVar(value=self.system_options[0]),
            finish=tk.StringVar(value=self.finish_options[0]),
            
            door_size=tk.StringVar(value=self.door_options[0]),
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

        self.main_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=self.fg_color)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.tab_view = ctk.CTkTabview(self.main_frame, fg_color=self.fg_color)
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
            self.update_status("Info: No projects found. Create a new one.", self.text_color)
            
        self.vars['system'].trace('w', lambda *args: self._reposition_widgets())

    # --- HELPER FUNCTIONS MOVED INSIDE THE CLASS ---

    def generate_excel_report(self, excel_path, elevations_json_path, extra_materials_json_path, **kwargs):
        print(f"Mocking Excel Report Generation for: {excel_path}")
        print(f"Elevation data to save: {kwargs}")

        existing_elevations = {}
        if os.path.exists(elevations_json_path):
            try:
                with open(elevations_json_path, 'r') as f:
                    existing_elevations = json.load(f)
            except (IOError, json.JSONDecodeError):
                pass

        if 'delete_elevation_type' in kwargs:
            elev_type = kwargs['delete_elevation_type']
            if elev_type in existing_elevations:
                del existing_elevations[elev_type]
                print(f"Deleted elevation '{elev_type}' from JSON.")
        elif 'elevation_type' in kwargs:
            elev_type = kwargs['elevation_type']
            kwargs.pop('delete_elevation_type', None)
            existing_elevations[elev_type] = kwargs
            print(f"Saved/Updated elevation '{elev_type}' in JSON.")
        
        with open(elevations_json_path, 'w') as f:
            json.dump(existing_elevations, f, indent=4)
            
    def calculate_yes45tu_quantities(self, *args, **kwargs):
        total = kwargs.get('total', 0)
        doors = kwargs.get('doors', [])
        
        output = {
            'total_elevations': total,
            'total_doors': sum(d['count'] for d in doors),
            'door_specs': doors
        }
        return output

    @staticmethod
    def calculate_rectangle_area(width, height):
        return width * height

    @staticmethod
    def calculate_perimeter(width, height):
        return 2 * (width + height)

    # --- END HELPER FUNCTIONS ---

    def create_project_tab_widgets(self):
        self.project_tab.grid_columnconfigure(0, weight=1)
        self.project_tab.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.project_tab, text="Project Management", font=self.title_font, text_color=self.text_color).grid(row=0, column=0, columnspan=2, pady=(10, 20), sticky="ew")

        create_frame = ctk.CTkFrame(self.project_tab, fg_color=self.accent_color, corner_radius=10)
        create_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        create_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(create_frame, text="Create New Project", font=self.header_font).grid(row=0, column=0, columnspan=2, pady=(10, 5))
        ctk.CTkLabel(create_frame, text="Name:", font=self.label_font).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        ctk.CTkEntry(create_frame, textvariable=self.vars['new_project_name'], fg_color=self.fg_color, font=self.label_font).grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        ctk.CTkButton(create_frame, text="Create", command=self.create_new_project, fg_color=self.fg_color, hover_color=self.accent_hover, font=self.button_font).grid(row=2, column=0, columnspan=2, pady=(5, 10))

        select_frame = ctk.CTkFrame(self.project_tab, fg_color=self.accent_color, corner_radius=10)
        select_frame.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")
        select_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(select_frame, text="Manage Existing Projects", font=self.header_font).grid(row=0, column=0, columnspan=2, pady=(10, 5))
        ctk.CTkLabel(select_frame, text="Select:", font=self.label_font).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.project_dropdown = ctk.CTkOptionMenu(select_frame, values=[], variable=self.vars['selected_project'], command=self.on_project_select, fg_color=self.fg_color, button_color=self.fg_color, button_hover_color=self.accent_hover, text_color=self.text_color, font=self.label_font)
        self.project_dropdown.grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        ctk.CTkButton(select_frame, text="Delete Selected", fg_color=self.error_color, hover_color="#8b1a1a", command=self.delete_current_project, font=self.button_font).grid(row=2, column=0, columnspan=2, pady=(5, 10))
        
        self.status_label = ctk.CTkLabel(self.project_tab, text="", text_color=self.text_color, font=self.label_font)
        self.status_label.grid(row=2, column=0, columnspan=2, pady=10)

    def create_elevation_tab_widgets(self):
        self.elevation_tab.grid_columnconfigure(0, weight=1, minsize=400)
        self.elevation_tab.grid_columnconfigure(1, weight=1, minsize=400)
        self.elevation_tab.grid_rowconfigure(0, weight=1)

        self.scroll_frame = ctk.CTkScrollableFrame(self.elevation_tab, fg_color=self.fg_color)
        self.scroll_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
        self.scroll_frame.grid_columnconfigure(0, weight=1)
        self.scroll_frame.grid_columnconfigure(1, weight=1)
        
        left_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        left_frame.grid_columnconfigure(1, weight=1)
        self.left_row = 0
        
        self._add_section_header(left_frame, "General Details")
        self._add_input_row(left_frame, "Select System:", ctk.CTkOptionMenu, self.vars['system'], self.system_options)
        self._add_input_row(left_frame, "Select Finish:", ctk.CTkOptionMenu, self.vars['finish'], self.finish_options)
        
        self._add_section_header(left_frame, "Elevation Specifications")
        self.saved_elevations_label, self.saved_elevations_option_menu = self._add_input_row(left_frame, "Saved Elevations:", ctk.CTkOptionMenu, self.vars['saved_elevation_types'], [], command=self.on_saved_elevation_select)
        self.elevation_type_label, self.elevation_type_entry = self._add_input_row(left_frame, "Elevation Type:", ctk.CTkEntry, self.vars['elevation_type'], None)
        self.total_count_label, self.total_count_entry = self._add_input_row(left_frame, "Total Count:", ctk.CTkEntry, self.vars['total_count'], None)
        self.bays_wide_label, self.bays_wide_entry = self._add_input_row(left_frame, "# Bays Wide:", ctk.CTkEntry, self.vars['bays_wide'], None)
        self.bays_tall_label, self.bays_tall_entry = self._add_input_row(left_frame, "# Bays Tall:", ctk.CTkEntry, self.vars['bays_tall'], None)
        self.opening_width_label, self.opening_width_entry = self._add_input_row(left_frame, "Opening Width (in):", ctk.CTkEntry, self.vars['opening_width'], None)
        self.opening_height_label, self.opening_height_entry = self._add_input_row(left_frame, "Opening Height (in):", ctk.CTkEntry, self.vars['opening_height'], None)
        
        button_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        button_frame.grid(row=self.left_row, column=0, columnspan=2, pady=(20, 10))
        ctk.CTkButton(button_frame, text="Save Elevation", command=self.save_elevation_type, font=self.button_font, fg_color=self.accent_color, hover_color=self.accent_hover).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Delete Elevation", fg_color=self.error_color, hover_color="#8b1a1a", command=self.delete_elevation_type, font=self.button_font).pack(side="left", padx=10)
        self.left_row += 1

        right_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        right_frame.grid_columnconfigure(0, weight=1)
        self.right_row = 0
        
        self._add_section_header(right_frame, "Door Management", right_row_counter=True)
        
        door_management_frame = ctk.CTkFrame(right_frame, fg_color=self.accent_color, corner_radius=10, border_color=self.border_color, border_width=1)
        door_management_frame.grid(row=self.right_row, column=0, sticky="ew", padx=10, pady=10)
        door_management_frame.grid_columnconfigure(1, weight=1)
        self.right_row += 1

        self._create_door_input_widgets(door_management_frame)

        self.door_listbox = tk.Listbox(door_management_frame, height=5, font=(self.font_family, 12), fg=self.text_color, bg=self.accent_color, selectbackground=self.accent_hover, selectforeground=self.text_color, borderwidth=0, highlightthickness=0)
        self.door_listbox.grid(row=self.door_input_row, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 10))
        self.door_listbox.bind('<<ListboxSelect>>', self.select_door_for_edit)
        
        door_action_buttons_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        door_action_buttons_frame.grid(row=self.right_row, column=0, sticky="ew", pady=(5, 10))
        ctk.CTkButton(door_action_buttons_frame, text="Add Door", command=self.add_door, font=self.button_font, fg_color=self.accent_color, hover_color=self.accent_hover).pack(side="left", expand=True, fill="x", padx=5)
        ctk.CTkButton(door_action_buttons_frame, text="Update", command=self.update_door, font=self.button_font, fg_color=self.accent_color, hover_color=self.accent_hover).pack(side="left", expand=True, fill="x", padx=5)
        ctk.CTkButton(door_action_buttons_frame, text="Delete", command=self.delete_door, font=self.button_font, fg_color=self.error_color, hover_color="#8b1a1a").pack(side="left", expand=True, fill="x", padx=5)
        self.right_row += 1

        self._reposition_widgets()

    def _add_section_header(self, parent, text, right_row_counter=False):
        """Helper to add a bold section header within a frame and increment the appropriate row counter."""
        row_attr_name = "right_row" if right_row_counter else "left_row"
        current_row = getattr(self, row_attr_name)
        
        lbl = ctk.CTkLabel(parent, text=text, font=self.header_font, text_color=self.text_color)
        lbl.grid(row=current_row, column=0, columnspan=2, sticky="w", pady=(20, 5), padx=5)
        
        setattr(self, row_attr_name, current_row + 1)
        
    def _add_input_row(self, parent, label_text, widget_type, variable, options=None, **kwargs):
        """Helper to create a label and input widget in a grid row, and increment the row counter."""
        current_row = self.left_row
        
        lbl = ctk.CTkLabel(parent, text=label_text, font=self.label_font, text_color=self.text_color)
        lbl.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        
        if widget_type == ctk.CTkEntry:
            widget = widget_type(parent, textvariable=variable, font=self.label_font, fg_color=self.accent_color, text_color=self.text_color, border_color=self.border_color, **kwargs)
        elif widget_type == ctk.CTkOptionMenu:
            widget = widget_type(parent, values=options if options else [""], variable=variable, font=self.label_font, fg_color=self.accent_color, button_color=self.accent_color, button_hover_color=self.accent_hover, text_color=self.text_color, **kwargs)
        
        widget.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5)
        
        widget_id = re.sub(r'[^a-z0-9_]', '', label_text.replace(" ", "_").lower())
        self.widgets[f"label_{widget_id}"] = lbl
        self.widgets[f"widget_{widget_id}"] = widget

        self.left_row += 1
        return lbl, widget

    def _create_door_input_widgets(self, parent_frame):
        self.door_input_row = 0
        
        self._add_input_row_to_frame(parent_frame, "Door Size:", ctk.CTkOptionMenu, self.vars['door_size'], self.door_options)
        self._add_input_row_to_frame(parent_frame, "Number of Doors:", ctk.CTkEntry, self.vars['door_count'])
        self._add_input_row_to_frame(parent_frame, "Stile Style:", ctk.CTkOptionMenu, self.vars['stile'], self.stile_options)
        
        hardware_label = ctk.CTkLabel(parent_frame, text="Select Hardware:", font=self.label_font, text_color=self.text_color)
        hardware_label.grid(row=self.door_input_row, column=0, sticky="w", padx=10, pady=5)
        self.door_input_row += 1

        self.hardware_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        self.hardware_frame.grid(row=self.door_input_row, column=0, columnspan=2, sticky="ew", padx=10)
        self.door_input_row += 1
        
        row_idx = 0
        for option in self.hardware_options:
            cb = ctk.CTkCheckBox(self.hardware_frame, text=option, variable=self.hardware_vars[option], fg_color=self.fg_color, hover_color=self.accent_hover, border_color=self.border_color, font=self.label_font, text_color=self.text_color)
            cb.grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
            row_idx += 1

    def _add_input_row_to_frame(self, parent, label_text, widget_type, variable, options=None, **kwargs):
        lbl = ctk.CTkLabel(parent, text=label_text, font=self.label_font, text_color=self.text_color)
        lbl.grid(row=self.door_input_row, column=0, sticky="w", padx=10, pady=5)
        
        if widget_type == ctk.CTkEntry:
            widget = widget_type(parent, textvariable=variable, font=self.label_font, fg_color=self.fg_color, text_color=self.text_color, border_color=self.border_color, **kwargs)
        elif widget_type == ctk.CTkOptionMenu:
            widget = widget_type(parent, values=options if options else [""], variable=variable, font=self.label_font, fg_color=self.fg_color, button_color=self.accent_color, button_hover_color=self.accent_hover, text_color=self.text_color, **kwargs)
        
        widget.grid(row=self.door_input_row, column=1, sticky="ew", padx=10, pady=5)
        self.door_input_row += 1
    
    def _reposition_widgets(self, *args):
        show_bays = self.vars['system'].get() == "YES 45TU FRONT SET(OG)"
        if show_bays:
            self.bays_wide_label.grid(row=self.bays_wide_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.bays_wide_entry.grid(row=self.bays_wide_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
            self.bays_tall_label.grid(row=self.bays_tall_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.bays_tall_entry.grid(row=self.bays_tall_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
        else:
            self.bays_wide_label.grid_forget()
            self.bays_wide_entry.grid_forget()
            self.bays_tall_label.grid_forget()
            self.bays_tall_entry.grid_forget()
            
    def add_door(self):
        if not self.current_project_name:
            self.update_status("Error: Please select a project first.", self.error_color)
            return
        
        try:
            door_size = self.vars['door_size'].get()
            if door_size == 'None':
                self.update_status("Error: Cannot add 'None' as a door.", self.error_color)
                return
            
            door_count_str = self.vars['door_count'].get()
            if not door_count_str:
                self.update_status("Error: Please enter a number of doors.", self.error_color)
                return
            door_count = int(door_count_str)
            
            stile_style = self.vars['stile'].get()
            hardware = [opt for opt, var in self.hardware_vars.items() if var.get()]
            
            new_door = {
                'size': door_size,
                'count': door_count,
                'stile': stile_style,
                'hardware': hardware
            }
            
            self.current_elevation_doors.append(new_door)
            self.update_door_listbox()
            self.clear_door_form()
            self.update_status("Door added successfully.", self.success_color)

        except ValueError:
            self.update_status("Error: Number of doors must be an integer.", self.error_color)

    def update_door(self):
        if self.selected_door_index is None:
            self.update_status("Error: No door selected to update.", self.error_color)
            return
        
        if not self.current_project_name:
            self.update_status("Error: Please select a project first.", self.error_color)
            return

        try:
            door_size = self.vars['door_size'].get()
            if door_size == 'None':
                self.update_status("Error: Cannot update to 'None' as a door.", self.error_color)
                return
            
            door_count_str = self.vars['door_count'].get()
            if not door_count_str:
                self.update_status("Error: Please enter a number of doors.", self.error_color)
                return
            door_count = int(door_count_str)

            stile_style = self.vars['stile'].get()
            hardware = [opt for opt, var in self.hardware_vars.items() if var.get()]
            
            updated_door = {
                'size': door_size,
                'count': door_count,
                'stile': stile_style,
                'hardware': hardware
            }
            
            self.current_elevation_doors[self.selected_door_index] = updated_door
            self.update_door_listbox()
            self.clear_door_form()
            self.selected_door_index = None
            self.update_status("Selected door updated successfully.", self.success_color)
        except ValueError:
            self.update_status("Error: Number of doors must be an integer.", self.error_color)

    def delete_door(self):
        if self.selected_door_index is None:
            self.update_status("Error: No door selected to delete.", self.error_color)
            return
        
        if not self.current_project_name:
            self.update_status("Error: Please select a project first.", self.error_color)
            return

        door_desc = self.door_listbox.get(self.selected_door_index)
        if not tk.messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete this door?\n{door_desc}"):
            self.update_status("Deletion cancelled.", self.text_color)
            return

        del self.current_elevation_doors[self.selected_door_index]
        self.update_door_listbox()
        self.clear_door_form()
        self.selected_door_index = None
        self.update_status("Door deleted successfully.", self.success_color)

    def update_door_listbox(self):
        self.door_listbox.delete(0, tk.END)
        for door in self.current_elevation_doors:
            hardware_str = ", ".join(door['hardware']) if door['hardware'] else "None"
            entry = f"{door['count']}x {door['size']} | Stile: {door['stile']} | Hardware: {hardware_str}"
            self.door_listbox.insert(tk.END, entry)

    def select_door_for_edit(self, event):
        try:
            index = self.door_listbox.curselection()[0]
            self.selected_door_index = index
            door = self.current_elevation_doors[index]
            
            self.vars['door_size'].set(door['size'])
            self.vars['door_count'].set(str(door['count']))
            self.vars['stile'].set(door['stile'])
            
            for var in self.hardware_vars.values():
                var.set(False)
                
            for hw in door['hardware']:
                if hw in self.hardware_vars:
                    self.hardware_vars[hw].set(True)
                    
            self.update_status(f"Selected door {index+1} for editing.", self.text_color)
        except IndexError:
            self.selected_door_index = None
            self.clear_door_form()

    def clear_door_form(self):
        self.vars['door_size'].set(self.door_options[0])
        self.vars['door_count'].set("")
        self.vars['stile'].set(self.stile_options[0])
        for var in self.hardware_vars.values():
            var.set(False)
        self.selected_door_index = None
        
    def on_saved_elevation_select(self, elev_type):
        if not elev_type or elev_type not in self.saved_elevations:
            self.clear_elevation_form()
            self.update_status("Info: Form cleared.", self.text_color)
            return
            
        data = self.saved_elevations[elev_type]
        self.clear_door_form()
        self.current_elevation_doors = data.get('doors', [])
        self.update_door_listbox()
        self.vars['elevation_type'].set(elev_type)
        self.vars['system'].set(str(data.get('system_input', self.system_options[0])))
        self.vars['finish'].set(str(data.get('finish_input', self.finish_options[0])))
        self.vars['total_count'].set(str(data.get('total_count', '')))
        self.vars['bays_wide'].set(str(data.get('bays_wide', '')))
        self.vars['bays_tall'].set(str(data.get('bays_tall', '')))
        self.vars['opening_width'].set(str(data.get('opening_width', '')))
        self.vars['opening_height'].set(str(data.get('opening_height', '')))
        self._reposition_widgets()
        self.update_status(f"Loaded elevation: {elev_type}", self.text_color)

    def clear_elevation_form(self):
        for var in ['elevation_type', 'total_count', 'bays_wide', 'bays_tall', 'opening_width', 'opening_height']:
            self.vars[var].set("")
        self.vars['system'].set(self.system_options[0])
        self.vars['finish'].set(self.finish_options[0])
        self.current_elevation_doors = []
        self.update_door_listbox()
        self.clear_door_form()
        self._reposition_widgets()
        
    def save_elevation_type(self):
        if not self.current_project_name:
            self.update_status("Error: Please create or select a project first.", self.error_color)
            return
        try:
            v = self.vars
            elev = v['elevation_type'].get().strip()
            
            if not elev:
                self.update_status("Error: Please enter an elevation type.", self.error_color)
                return

            is_update = elev in self.saved_elevations
            system = v['system'].get()
            finish = v['finish'].get()
            total_str = v['total_count'].get()
            total = int(total_str) if total_str else 0
            ow_str = v['opening_width'].get()
            ow = float(ow_str) if ow_str else 0.0
            oh_str = v['opening_height'].get()
            oh = float(oh_str) if oh_str else 0.0
            bays_wide_str = v['bays_wide'].get()
            bays_wide = int(bays_wide_str) if system == self.system_options[0] and bays_wide_str else 0
            bays_tall_str = v['bays_tall'].get()
            bays_tall = int(bays_tall_str) if system == self.system_options[0] and bays_tall_str else 0

            calculated = []
            if system == self.system_options[0]:
                calculated = self.calculate_yes45tu_quantities(
                    bays_wide=bays_wide, bays_tall=bays_tall, total=total, 
                    opening_width=ow, opening_height=oh, doors=self.current_elevation_doors
                )
                
            sqft_per = self.calculate_rectangle_area(ow / 12, oh / 12) if ow and oh else 0
            total_sqft = sqft_per * total
            perimeter = self.calculate_perimeter(ow / 12, oh / 12) if ow and oh else 0
            total_perimeter = perimeter * total

            self.generate_excel_report(
                excel_path=self.current_excel_path,
                elevations_json_path=self.current_elevations_json_path,
                extra_materials_json_path=self.current_extra_materials_json_path,
                system_input=system, finish_input=finish, elevation_type=elev,
                total_count=total, bays_wide=bays_wide, bays_tall=bays_tall,
                opening_width=ow, opening_height=oh, sqft_per_type=sqft_per,
                total_sqft=total_sqft, perimeter_ft=perimeter, total_perimeter_ft=total_perimeter,
                calculated_outputs=calculated, doors=self.current_elevation_doors,
                mode="save_or_update"
            )
            
            self.load_saved_elevations_for_current_project()
            self.vars['saved_elevation_types'].set(elev)
            
            if is_update:
                self.update_status(f"Elevation '{elev}' updated successfully.", self.success_color)
            else:
                self.update_status(f"Elevation '{elev}' saved successfully.", self.success_color)

        except ValueError as e:
            self.update_status(f"Error: Invalid input format. Please check your numbers: {e}", self.error_color)
        except Exception as e:
            self.update_status(f"An unexpected error occurred: {e}", self.error_color)

    def delete_elevation_type(self):
        if not self.current_project_name:
            self.update_status("Error: Please select a project first.", self.error_color)
            return
        elev = self.vars['saved_elevation_types'].get()
        if not elev:
            self.update_status("Error: No elevation selected to delete.", self.error_color)
            return
            
        if not tk.messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete elevation '{elev}'?"):
            self.update_status("Deletion cancelled.", self.text_color)
            return
            
        self.generate_excel_report(
            excel_path=self.current_excel_path,
            elevations_json_path=self.current_elevations_json_path,
            extra_materials_json_path=self.current_extra_materials_json_path,
            delete_elevation_type=elev
        )
        self.load_saved_elevations_for_current_project()
        self.clear_elevation_form()
        self.update_status(f"Elevation '{elev}' deleted successfully.", self.success_color)

    def load_project_list(self):
        if os.path.exists(MASTER_PROJECT_LIST_FILE):
            try:
                with open(MASTER_PROJECT_LIST_FILE, 'r') as f:
                    self.all_projects = json.load(f)
            except (IOError, json.JSONDecodeError):
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
            self.update_status("Error: Please enter a name for the new project.", self.error_color)
            return
        if new_name in self.all_projects:
            self.update_status(f"Error: Project '{new_name}' already exists.", self.error_color)
            return
        
        self.all_projects.append(new_name)
        self.save_project_list()
        self.update_project_dropdown()
        self.vars['selected_project'].set(new_name)
        self.on_project_select(new_name)
        self.update_status(f"Project '{new_name}' created.", self.success_color)
        self.vars['new_project_name'].set("")
        self.tab_view.set("Elevation Details")

    def on_project_select(self, project_name):
        self.current_project_name = project_name
        self.set_current_project_paths()
        self.load_saved_elevations_for_current_project()
        self.clear_elevation_form()
        self.update_status(f"Switched to project '{project_name}'", self.text_color)

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
            except (IOError, json.JSONDecodeError):
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
            self.update_status("Error: No project selected to delete.", self.error_color)
            return
        
        if not tk.messagebox.askyesno(
            "Confirm Deletion", f"Are you sure you want to delete project '{self.current_project_name}' and all its data?"
        ):
            self.update_status("Project deletion cancelled.", self.text_color)
            return
            
        try:
            files_to_delete = [self.current_excel_path, self.current_elevations_json_path, self.current_extra_materials_json_path]
            for file_path in files_to_delete:
                if os.path.exists(file_path): os.remove(file_path)
            
            if self.current_project_name in self.all_projects:
                self.all_projects.remove(self.current_project_name)
                self.save_project_list()
                self.update_project_dropdown()
            
            self.update_status(f"Project '{self.current_project_name}' and its files deleted.", self.success_color)
            if self.all_projects:
                self.on_project_select(self.all_projects[0])
            else:
                self.vars['selected_project'].set("")
                self.current_project_name = ""
                self.clear_elevation_form()
                self.saved_elevations = {}
                self.update_saved_elevation_dropdown()
                self.update_status("Info: No projects remaining. Create a new one.", self.text_color)
        except Exception as e:
            self.update_status(f"Error: Could not delete project '{self.current_project_name}': {e}", self.error_color)
    
    def update_status(self, message, color=None):
        if color is None:
            color = self.text_color
        self.status_label.configure(text=message, text_color=color)

if __name__ == "__main__":
    app = App()
    app.after(10, lambda: app.state('zoomed'))
    app.mainloop()