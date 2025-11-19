import customtkinter as ctk
import tkinter as tk
import tkinter.messagebox
import os

# Refactored utility modules
from utils.excel_generator import generate_excel_report
from utils.formulas import calculate_total_door_area, calculate_glass_to_add_back
from utils.path_manager import (
    PROJECTS_DIR,
    get_project_paths,
    get_door_json_path,
    get_unique_report_path
)
from utils.file_operations import (
    load_project_list,
    save_project_list,
    load_elevations,
    save_elevations,
    load_doors,
    save_doors,
    ensure_json_file_exists,
    create_excel_file,
    load_elevation_data
)
from utils.project_manager import (
    create_project,
    delete_project,
    load_projects,
    initialize_project_files
)
from utils.elevation_manager import (
    transform_elevation_data_to_ui,
    build_elevation_data,
    save_elevation,
    delete_elevation,
    get_elevation_names
)
from utils.door_manager import (
    get_door_file_path,
    ensure_door_file,
    load_doors_for_elevation,
    save_doors_for_elevation,
    format_door_for_display
)
from utils.validation import (
    parse_custom_bays,
    validate_door_addition,
    validate_door_update,
    validate_door_count
)

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
        self.current_door_json_path = ""
        
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
            selected_project=tk.StringVar(),
            custom_bay_widths=tk.StringVar(),  # New variable for custom bay widths
            custom_bay_heights=tk.StringVar()  # New variable for custom bay heights
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
        
        self.on_system_change(self.vars['system'].get())

        self.load_project_list()
        if self.all_projects:
            self.vars['selected_project'].set(self.all_projects[0])
            self.on_project_select(self.all_projects[0])
        else:
            self.update_project_dropdown()
            self.update_status("Info: No projects found. Create a new one.", self.text_color)
            
        self.vars['system'].trace('w', lambda *args: self.on_system_change(self.vars['system'].get()))

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
        self.custom_bay_widths_label, self.custom_bay_widths_entry = self._add_input_row(left_frame, "Custom Bay Widths (comma-separated):", ctk.CTkEntry, self.vars['custom_bay_widths'], None)
        self.custom_bay_heights_label, self.custom_bay_heights_entry = self._add_input_row(left_frame, "Custom Bay Heights (comma-separated):", ctk.CTkEntry, self.vars['custom_bay_heights'], None)
        
        button_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        button_frame.grid(row=self.left_row, column=0, columnspan=2, pady=(20, 10))
        ctk.CTkButton(button_frame, text="Save Elevation", command=self.save_elevation_type, font=self.button_font, fg_color=self.accent_color, hover_color=self.accent_hover).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Delete Elevation", fg_color=self.error_color, hover_color="#8b1a1a", command=self.delete_elevation_type, font=self.button_font).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Generate Report", command=self.generate_unique_report, font=self.button_font, fg_color=self.accent_color, hover_color=self.accent_hover).pack(side="left", padx=10)
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

        self.door_listbox = tk.Listbox(door_management_frame, height=5, font=(self.font_family, 12), fg="white", bg=self.accent_color, selectbackground=self.accent_hover, selectforeground="white", borderwidth=0, highlightthickness=0)
        self.door_listbox.grid(row=self.door_input_row, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 10))
        self.door_listbox.bind('<<ListboxSelect>>', self.select_door_for_edit)
        
        door_action_buttons_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        door_action_buttons_frame.grid(row=self.right_row, column=0, sticky="ew", pady=(5, 10))
        ctk.CTkButton(door_action_buttons_frame, text="Add Door", command=self.add_door, font=self.button_font, fg_color=self.accent_color, hover_color=self.accent_hover).pack(side="left", expand=True, fill="x", padx=5)
        ctk.CTkButton(door_action_buttons_frame, text="Update", command=self.update_door, font=self.button_font, fg_color=self.accent_color, hover_color=self.accent_hover).pack(side="left", expand=True, fill="x", padx=5)
        ctk.CTkButton(door_action_buttons_frame, text="Delete", command=self.delete_door, font=self.button_font, fg_color=self.error_color, hover_color="#8b1a1a").pack(side="left", expand=True, fill="x", padx=5)
        self.right_row += 1
        
        self.elevation_status_label = ctk.CTkLabel(self.elevation_tab, text="", text_color=self.text_color, font=self.label_font)
        self.elevation_status_label.grid(row=1, column=0, columnspan=2, pady=10)

    def _add_section_header(self, parent, text, right_row_counter=False):
        row_attr_name = "right_row" if right_row_counter else "left_row"
        current_row = getattr(self, row_attr_name)
        
        lbl = ctk.CTkLabel(parent, text=text, font=self.header_font, text_color=self.text_color)
        lbl.grid(row=current_row, column=0, columnspan=2, sticky="w", pady=(20, 5), padx=5)
        
        setattr(self, row_attr_name, current_row + 1)
        
    def _add_input_row(self, parent, label_text, widget_type, variable, options=None, **kwargs):
        current_row = self.left_row
        
        lbl = ctk.CTkLabel(parent, text=label_text, font=self.label_font, text_color=self.text_color)
        lbl.grid(row=current_row, column=0, sticky="w", padx=10, pady=5)
        
        if widget_type == ctk.CTkEntry:
            widget = widget_type(parent, textvariable=variable, font=self.label_font, fg_color=self.accent_color, text_color=self.text_color, border_color=self.border_color, **kwargs)
        elif widget_type == ctk.CTkOptionMenu:
            widget = widget_type(parent, values=options if options else [""], variable=variable, font=self.label_font, fg_color=self.accent_color, button_color=self.accent_color, button_hover_color=self.accent_hover, text_color=self.text_color, **kwargs)
        
        widget.grid(row=current_row, column=1, sticky="ew", padx=10, pady=5)
        
        widget_id = label_text.replace(" ", "_").replace(":", "").lower()
        self.widgets[f"label_{widget_id}"] = lbl
        self.widgets[f"widget_{widget_id}"] = widget

        self.left_row += 1
        return lbl, widget

    def _create_door_input_widgets(self, parent_frame):
        self.door_input_row = 0
        
        self._add_input_row_to_frame(parent_frame, "Door Size:", ctk.CTkOptionMenu, self.vars['door_size'], self.door_options)
        self._add_input_row_to_frame(parent_frame, "Number of Doors:", ctk.CTkEntry, self.vars['door_count'])
        self._add_input_row_to_frame(parent_frame, "Style:", ctk.CTkOptionMenu, self.vars['stile'], self.stile_options)
        
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
    
    def on_system_change(self, *args):
        show_bays = self.vars['system'].get() == "YES 45TU FRONT SET(OG)"
        if show_bays:
            self.bays_wide_label.grid(row=self.bays_wide_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.bays_wide_entry.grid(row=self.bays_wide_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
            self.bays_tall_label.grid(row=self.bays_tall_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.bays_tall_entry.grid(row=self.bays_tall_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
            self.opening_width_label.grid(row=self.opening_width_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.opening_width_entry.grid(row=self.opening_width_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
            self.opening_height_label.grid(row=self.opening_height_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.opening_height_entry.grid(row=self.opening_height_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
            self.custom_bay_widths_label.grid(row=self.custom_bay_widths_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.custom_bay_widths_entry.grid(row=self.custom_bay_widths_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
            self.custom_bay_heights_label.grid(row=self.custom_bay_heights_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.custom_bay_heights_entry.grid(row=self.custom_bay_heights_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
        else:
            self.bays_wide_label.grid_forget()
            self.bays_wide_entry.grid_forget()
            self.bays_tall_label.grid_forget()
            self.bays_tall_entry.grid_forget()
            self.custom_bay_widths_label.grid_forget()
            self.custom_bay_widths_entry.grid_forget()
            self.custom_bay_heights_label.grid_forget()
            self.custom_bay_heights_entry.grid_forget()
            self.opening_width_label.grid(row=self.opening_width_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.opening_width_entry.grid(row=self.opening_width_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
            self.opening_height_label.grid(row=self.opening_height_entry.grid_info()['row'], column=0, sticky="w", padx=10, pady=5)
            self.opening_height_entry.grid(row=self.opening_height_entry.grid_info()['row'], column=1, sticky="ew", padx=10, pady=5)
            
    def load_project_list(self):
        try:
            self.all_projects = load_projects()
            self.update_project_dropdown()
        except Exception as e:
            self.update_status(f"Error loading project list: {e}", self.error_color)
            self.all_projects = []
            self.update_project_dropdown()

    def save_project_list(self):
        try:
            save_project_list(self.all_projects)
        except Exception as e:
            self.update_status(f"Error saving project list: {e}", self.error_color)

    def update_project_dropdown(self):
        self.project_dropdown.configure(values=self.all_projects)
        if self.all_projects and self.current_project_name in self.all_projects:
            self.vars['selected_project'].set(self.current_project_name)
        elif self.all_projects:
            self.vars['selected_project'].set(self.all_projects[0])
            self.current_project_name = self.all_projects[0]
            self.set_current_project_paths()
        else:
            self.vars['selected_project'].set("")

    def create_new_project(self):
        new_name = self.vars['new_project_name'].get().strip()
        success, error_msg, updated_projects = create_project(new_name, self.all_projects)
        
        if not success:
            self.update_status(f"Error: {error_msg}", self.error_color)
            return
        
        self.all_projects = updated_projects
        self.update_project_dropdown()
        self.vars['selected_project'].set(new_name)
        self.on_project_select(new_name)
        self.update_status(f"New project '{new_name}' created.", self.success_color)
        self.vars['new_project_name'].set("")

    def on_project_select(self, project_name):
        self.current_project_name = project_name
        self.set_current_project_paths()
        self.load_saved_elevations_for_current_project()
        self.clear_form()
        self.update_status(f"Switched to project '{project_name}'.", self.text_color)
        self.tab_view.set("Elevation Details")

    def set_current_project_paths(self):
        paths = get_project_paths(self.current_project_name)
        self.current_excel_path = paths['excel_path']
        self.current_elevations_json_path = paths['elevations_json_path']
        self.current_extra_materials_json_path = paths['extra_materials_json_path']
        
        # Initialize project files if they don't exist
        if self.current_project_name:
            initialize_project_files(self.current_project_name)

    def load_saved_elevations_for_current_project(self):
        try:
            self.saved_elevations = load_elevations(self.current_elevations_json_path)
            
            self.update_saved_elevation_dropdown()
            if self.saved_elevations:
                first_elevation = sorted(self.saved_elevations.keys())[0]
                self.vars['saved_elevation_types'].set(first_elevation)
                self.on_saved_elevation_select(first_elevation)
            
            self.update_status(f"Loaded elevations for '{self.current_project_name}'.", self.success_color)
        except Exception as e:
            self.update_status(f"Error loading elevations: {e}", self.error_color)
            self.saved_elevations = {}
            self.update_saved_elevation_dropdown()
            self.clear_form()
            ensure_json_file_exists(self.current_elevations_json_path, default_data={})

    def delete_current_project(self):
        if not self.current_project_name:
            self.update_status("Error: No project selected to delete.", self.error_color)
            return

        if not tkinter.messagebox.askyesno(
            "Confirm Deletion",
            f"Are you sure you want to delete project '{self.current_project_name}' and all its data? This cannot be undone."
        ):
            self.update_status("Project deletion cancelled.", self.text_color)
            return

        success, error_msg, updated_projects, next_project = delete_project(
            self.current_project_name, self.all_projects
        )
        
        if not success:
            self.update_status(f"Error: {error_msg}", self.error_color)
            return

        self.all_projects = updated_projects
        self.update_project_dropdown()
        self.update_status(f"Project '{self.current_project_name}' deleted.", self.success_color)
        
        if next_project:
            self.vars['selected_project'].set(next_project)
            self.on_project_select(next_project)
        else:
            self.vars['selected_project'].set("")
            self.current_project_name = ""
            self.clear_form()
            self.saved_elevations = {}
            self.update_saved_elevation_dropdown()
            self.update_status("Info: No projects remaining. Create a new project.", self.text_color)

    def on_saved_elevation_select(self, elev_type):
        self.door_listbox.delete(0, tk.END)
        self.clear_door_form()
        
        if not elev_type or elev_type not in self.saved_elevations:
            self.clear_form()
            return
            
        data = self.saved_elevations[elev_type]
        ui_data = transform_elevation_data_to_ui(data)
        
        for var_key, value in ui_data.items():
            if var_key in self.vars:
                self.vars[var_key].set(value)
                
        self.vars['elevation_type'].set(elev_type)
        self.on_system_change(self.vars['system'].get())

        self.current_door_json_path = self._ensure_door_file(elev_type)
        self.update_door_listbox()
        
        self.update_status(f"Elevation '{elev_type}' loaded.", self.success_color)

    def _get_door_json_path(self, elev_type):
        return get_door_file_path(self.current_project_name, elev_type)

    def _ensure_door_file(self, elev_type):
        """Ensure the door JSON file for this elevation exists (creates empty list if missing)."""
        return ensure_door_file(self.current_project_name, elev_type)

    def update_door_listbox(self):
        self.door_listbox.delete(0, tk.END)
        if not self.current_door_json_path:
            return []
        
        try:
            doors = load_doors(self.current_door_json_path)
            for i, door in enumerate(doors):
                door_text = format_door_for_display(door, i)
                self.door_listbox.insert(tk.END, door_text)
            return doors
        except Exception:
            return []

    def save_door_data(self, doors):
        if not self.current_door_json_path:
            self.update_status("Error: No elevation selected to save doors.", self.error_color)
            return

        try:
            save_doors(self.current_door_json_path, doors)
        except Exception as e:
            self.update_status(f"Error saving door data: {e}", self.error_color)
    

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

            system = v['system'].get()
            finish = v['finish'].get()
            total_count = int(v['total_count'].get())
            opening_width = float(v['opening_width'].get())
            opening_height = float(v['opening_height'].get())

            # Ensure we read doors for the elevation being saved, and ensure file exists
            self.current_door_json_path = self._ensure_door_file(elev)
            doors = self.update_door_listbox()
            
            # Build elevation data using refactored function
            elevation_data = build_elevation_data(
                system=system,
                finish=finish,
                total_count=total_count,
                opening_width=opening_width,
                opening_height=opening_height,
                bays_wide=int(v['bays_wide'].get()) if v['bays_wide'].get() else None,
                bays_tall=int(v['bays_tall'].get()) if v['bays_tall'].get() else None,
                custom_bay_widths_str=v['custom_bay_widths'].get(),
                custom_bay_heights_str=v['custom_bay_heights'].get(),
                doors=doors,
                system_options=self.system_options
            )

            # Save elevation
            save_elevation(self.current_elevations_json_path, elev, elevation_data)
            
            # Reload elevations to keep in sync
            self.saved_elevations = load_elevations(self.current_elevations_json_path)
            self.current_door_json_path = self._get_door_json_path(elev)

            # Generate Excel report
            generate_excel_report(
                excel_path=self.current_excel_path,
                elevations_json_path=self.current_elevations_json_path,
                extra_materials_json_path=self.current_extra_materials_json_path,
                system_input=system,
                finish_input=finish,
                elevation_type=elev,
                total_count=total_count,
                bays_wide=elevation_data.get('bays_wide', 0),
                bays_tall=elevation_data.get('bays_tall', 0),
                opening_width=opening_width,
                opening_height=opening_height,
                sqft_per_type=elevation_data['sqft_per_type'],
                total_sqft=elevation_data['total_sqft'],
                perimeter_ft=elevation_data['perimeter_ft'],
                total_perimeter_ft=elevation_data['total_perimeter_ft'],
                calculated_outputs=elevation_data['calculated_outputs'],
                completion_callback=None,
                doors=doors,
                custom_bay_widths=elevation_data.get('custom_bay_widths', []),
                custom_bay_heights=elevation_data.get('custom_bay_heights', [])
            )

            self.update_saved_elevation_dropdown()
            self.vars['saved_elevation_types'].set(elev)
            self.update_status(f"Elevation '{elev}' saved successfully.", self.success_color)

        except ValueError as e:
            self.update_status(f"Error: {e}", self.error_color)
        except Exception as e:
            self.update_status(f"An unexpected error occurred: {e}", self.error_color)

    def delete_elevation_type(self):
        if not self.current_project_name:
            self.update_status("Error: Please select a project first.", self.error_color)
            return

        elev = self.vars['saved_elevation_types'].get().strip()
        if elev:
            if delete_elevation(self.current_elevations_json_path, elev):
                # Reload elevations
                self.saved_elevations = load_elevations(self.current_elevations_json_path)
                
                # Delete door file
                door_file_path = self._get_door_json_path(elev)
                if door_file_path and os.path.exists(door_file_path):
                    os.remove(door_file_path)

                self.update_saved_elevation_dropdown()
                self.clear_form()
                self.update_status(f"Elevation '{elev}' and its doors deleted successfully.", self.success_color)

    def update_saved_elevation_dropdown(self):
        elevations = get_elevation_names(self.current_elevations_json_path) if self.current_elevations_json_path else []
        self.saved_elevations_option_menu.configure(values=elevations if elevations else [""])
        if elevations:
            self.vars['saved_elevation_types'].set(elevations[0])
        else:
            self.vars['saved_elevation_types'].set("")
            
    def clear_form(self):
        self.vars['elevation_type'].set("")
        self.vars['total_count'].set("")
        self.vars['bays_wide'].set("")
        self.vars['bays_tall'].set("")
        self.vars['custom_bay_widths'].set("")
        self.vars['custom_bay_heights'].set("")
        self.vars['opening_width'].set("")
        self.vars['opening_height'].set("")
        self.vars['system'].set(self.system_options[0])
        self.vars['finish'].set(self.finish_options[0])
        self.door_listbox.delete(0, tk.END)
        self.clear_door_form()

    def clear_door_form(self):
        self.vars['door_size'].set(self.door_options[0])
        self.vars['door_count'].set("")
        self.vars['stile'].set(self.stile_options[0])
        self.selected_door_index = None
        for var in self.hardware_vars.values():
            var.set(False)

    def add_door(self):
        door_size = self.vars['door_size'].get()
        door_count_str = self.vars['door_count'].get()
        stile = self.vars['stile'].get()
        hardware = {opt: self.hardware_vars[opt].get() for opt in self.hardware_options}

        if door_size == "None" or not door_count_str:
            self.update_status("Error: Please select a door size and count.", self.error_color)
            return

        elevation_name = self.vars['elevation_type'].get().strip()
        if not elevation_name:
            self.update_status("Error: Please enter an elevation type before adding doors.", self.error_color)
            return

        # Validate door count
        is_valid, error_msg, door_count = validate_door_count(door_count_str)
        if not is_valid:
            self.update_status(f"Error: {error_msg}", self.error_color)
            return

        self.current_door_json_path = self._ensure_door_file(elevation_name)
        doors = self.update_door_listbox()

        new_door = {'size': door_size, 'count': door_count, 'stile': stile, 'hardware': hardware}

        try:
            elevation_data = load_elevation_data(self.current_elevations_json_path, elevation_name)
            if not elevation_data:
                raise ValueError(f"Elevation '{elevation_name}' not found in saved data.")

            glass_area = elevation_data.get('total_sqft', 0.0)

            # Validate door addition
            is_valid, error_msg, _ = validate_door_addition(glass_area, doors, new_door)
            if not is_valid:
                self.update_status(f"Error: {error_msg}", self.error_color)
                return

        except Exception as e:
            self.update_status(f"Error checking area: {e}", self.error_color)
            return

        simulated_doors = doors + [new_door]
        self.save_door_data(simulated_doors)
        self.update_door_listbox()
        self.clear_door_form()
        self.save_elevation_type()
        self.update_status("Door added to the current elevation.", self.success_color)

    def update_door(self):
        if self.selected_door_index is None:
            self.update_status("Error: No door selected to update.", self.error_color)
            return

        door_size = self.vars['door_size'].get()
        door_count_str = self.vars['door_count'].get()
        stile = self.vars['stile'].get()
        hardware = {opt: self.hardware_vars[opt].get() for opt in self.hardware_options}

        # Validate door count
        is_valid, error_msg, door_count = validate_door_count(door_count_str)
        if not is_valid:
            self.update_status(f"Error: {error_msg}", self.error_color)
            return

        elevation_name = self.vars['elevation_type'].get().strip()
        if not elevation_name:
            self.update_status("Error: Please enter an elevation type first.", self.error_color)
            return

        doors = self.update_door_listbox()

        updated_door = {'size': door_size, 'count': door_count, 'stile': stile, 'hardware': hardware}

        simulated_doors = doors.copy()
        simulated_doors[self.selected_door_index] = updated_door

        try:
            elevation_data = load_elevation_data(self.current_elevations_json_path, elevation_name)
            if not elevation_data:
                raise ValueError(f"Elevation '{elevation_name}' not found in saved data.")

            glass_area = elevation_data.get('total_sqft', 0.0)

            # Validate door update
            is_valid, error_msg, _ = validate_door_update(glass_area, doors, self.selected_door_index, updated_door)
            if not is_valid:
                self.update_status(f"Error: {error_msg}", self.error_color)
                return

        except Exception as e:
            self.update_status(f"Error checking area: {e}", self.error_color)
            return

        door_num = self.selected_door_index + 1
        self.save_door_data(simulated_doors)
        self.update_door_listbox()
        self.clear_door_form()
        self.save_elevation_type()
        self.update_status(f"Door {door_num} updated.", self.success_color)

    def delete_door(self):
        if self.selected_door_index is None:
            self.update_status("Error: No door selected to delete.", self.error_color)
            return

        doors = self.update_door_listbox()

        door_num = self.selected_door_index + 1
        del doors[self.selected_door_index]

        self.save_door_data(doors)

        self.update_door_listbox()
        self.clear_door_form()
        self.save_elevation_type()
        self.update_status(f"Door {door_num} deleted.", self.success_color)
        self.selected_door_index = None

    def select_door_for_edit(self, event):
        selected_index = self.door_listbox.curselection()
        if selected_index:
            self.selected_door_index = selected_index[0]
            
            doors = self.update_door_listbox()
            if doors and self.selected_door_index < len(doors):
                door_data = doors[self.selected_door_index]
                self.vars['door_size'].set(door_data['size'])
                self.vars['door_count'].set(str(door_data['count']))
                self.vars['stile'].set(door_data['stile'])
                
                for opt, var in self.hardware_vars.items():
                    var.set(door_data['hardware'].get(opt, False))

    def update_status(self, message, color):
        self.elevation_status_label.configure(text=message, text_color=color)

    def generate_unique_report(self):
        if not self.current_project_name:
            self.update_status("Error: No project selected.", self.error_color)
            return
            
        unique_excel_path = get_unique_report_path(self.current_project_name)

        try:
            generate_excel_report(
                excel_path=unique_excel_path,
                elevations_json_path=self.current_elevations_json_path,
                extra_materials_json_path=self.current_extra_materials_json_path,
                system_input="",
                finish_input="",
                elevation_type="",
                total_count=0,
                bays_wide=0,
                bays_tall=0,
                opening_width=0,
                opening_height=0,
                sqft_per_type=0,
                total_sqft=0,
                perimeter_ft=0,
                total_perimeter_ft=0,
                calculated_outputs=[],
                completion_callback=None,
                reset=False,
                delete_elevation_type=None,
                doors=None,
                mode="export_all",
                custom_bay_widths=[],
                custom_bay_heights=[]
            )
            self.update_status(f"Report generated at '{unique_excel_path}'", self.success_color)
        except Exception as e:
            self.update_status(f"Error generating unique report: {e}", self.error_color)

if __name__ == "__main__":
    app = App()
    app.after(10, lambda: app.state('zoomed'))
    app.mainloop()