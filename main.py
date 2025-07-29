import customtkinter as ctk
import tkinter as tk
import json
import os
from openpyxl import Workbook # Import Workbook to create new Excel files

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

        # Ensure the projects directory exists
        os.makedirs(PROJECTS_DIR, exist_ok=True)

        # Constants
        self.system_options = ["YES 45TU FRONT SET(OG)", "Other"]
        self.finish_options = ["Clear", "Black", "Paint"]
        self.door_options = ['None', "3' X 7'", "3' X 8'", "3' X 9'", "6' X 7'", "6' X 8'", "6' X 9'"]
        self.saved_elevations = {} # Elevations for the *current* project
        self.all_projects = [] # To store names of all projects
        self.current_project_name = "" # Currently active project
        self.current_excel_path = ""
        self.current_elevations_json_path = ""
        self.current_extra_materials_json_path = ""


        # Tk variables
        vars_ = dict(
            system=tk.StringVar(value=self.system_options[0]),
            finish=tk.StringVar(value=self.finish_options[0]),
            door=tk.StringVar(value=self.door_options[0]),
            elevation_type=tk.StringVar(),
            total_count=tk.StringVar(),
            bays_wide=tk.StringVar(),
            bays_tall=tk.StringVar(),
            opening_width=tk.StringVar(),
            opening_height=tk.StringVar(),
            saved_elevation_types=tk.StringVar(),
            new_project_name=tk.StringVar(), # New variable for new project name
            selected_project=tk.StringVar() # New variable for selected project dropdown
        )
        self.vars = vars_

        # UI setup
        self.main_frame = ctk.CTkFrame(self, corner_radius=20)
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=30)
        self.main_frame.grid_columnconfigure(1, weight=1)

        # --- Project Management Section ---
        project_labels = [
            ("New Project Name:", 0),
            ("Select Project:", 1),
        ]
        self.widgets = {} # Store widgets for easy access
        for text, row in project_labels:
            lbl = ctk.CTkLabel(self.main_frame, text=text)
            lbl.grid(row=row, column=0, sticky="w", pady=5)
            self.widgets[f"label_project_{row}"] = lbl # Unique key for project labels

        self.new_project_entry = ctk.CTkEntry(self.main_frame, textvariable=vars_['new_project_name'])
        self.new_project_entry.grid(row=0, column=1, sticky="ew", pady=5)

        self.create_project_button = ctk.CTkButton(self.main_frame, text="Create New Project", command=self.create_new_project)
        self.create_project_button.grid(row=0, column=2, padx=10, pady=5)

        self.project_dropdown = ctk.CTkOptionMenu(
            self.main_frame,
            values=[], # Will be populated on load
            variable=vars_['selected_project'],
            command=self.on_project_select
        )
        self.project_dropdown.grid(row=1, column=1, sticky="ew", pady=5)
        # --- End Project Management Section ---

        # Adjust row numbers for existing UI elements (shifted down by 2 rows)
        labels = [
            ("Select System:", 2),
            ("Select Finish:", 3),
            ("Select Door Size:", 4),
            ("Saved Elevation Types:", 5),
            ("Elevation Type:", 6),
            ("Total Count:", 7),
            ("# Bays Wide:", 8),
            ("# Bays Tall:", 9),
            ("Opening Width (in inches):", 10),
            ("Opening Height (in inches):", 11),
        ]
        for text, row in labels:
            lbl = ctk.CTkLabel(self.main_frame, text=text)
            lbl.grid(row=row, column=0, sticky="w", pady=5)
            self.widgets[f"label_{row}"] = lbl # Update widget key for new row numbers

        self.system_dropdown = ctk.CTkOptionMenu(
            self.main_frame,
            values=self.system_options,
            variable=vars_['system'],
            command=self.on_system_change
        )
        self.system_dropdown.grid(row=2, column=1, sticky="ew", pady=5)

        self.finish_dropdown = ctk.CTkOptionMenu(
            self.main_frame, values=self.finish_options, variable=vars_['finish']
        )
        self.finish_dropdown.grid(row=3, column=1, sticky="ew", pady=5)

        self.door_dropdown = ctk.CTkOptionMenu(
            self.main_frame, values=self.door_options, variable=vars_['door']
        )
        self.door_dropdown.grid(row=4, column=1, sticky="ew", pady=5)

        self.saved_elevation_dropdown = ctk.CTkOptionMenu(
            self.main_frame,
            values=[],
            variable=vars_['saved_elevation_types'],
            command=self.on_saved_elevation_select
        )
        self.saved_elevation_dropdown.grid(row=5, column=1, sticky="ew", pady=5)

        entry_fields = ['elevation_type', 'total_count', 'bays_wide', 'bays_tall', 'opening_width', 'opening_height']
        for idx, field in enumerate(entry_fields, start=6): # Adjusted start row
            entry = ctk.CTkEntry(self.main_frame, textvariable=vars_[field])
            entry.grid(row=idx, column=1, sticky="ew", pady=5)
            self.widgets[f"entry_{field}"] = entry

        btn_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        btn_frame.grid(row=12, column=0, columnspan=3, sticky="e", pady=20, padx=(0, 30)) # Adjusted row and columnspan
        self.submit_button = ctk.CTkButton(btn_frame, text="Save Elevation Type", command=self.save_elevation_type)
        self.submit_button.pack(side="left", padx=(0, 20))
        self.delete_button = ctk.CTkButton(btn_frame, text="Delete Elevation", command=self.delete_elevation_type)
        self.delete_button.pack(side="left", padx=(0, 20))
        # Changed button from "Export All to Excel" to "Delete Project"
        self.delete_project_button = ctk.CTkButton(btn_frame, text="Delete Project", command=self.delete_current_project)
        self.delete_project_button.pack(side="left", padx=(0, 20))

        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="red")
        self.status_label.grid(row=13, column=0, columnspan=3) # Adjusted row and columnspan

        # Initialize project system
        self.load_project_list()
        if self.all_projects:
            # Select the first project by default or last active one
            self.vars['selected_project'].set(self.all_projects[0])
            self.on_project_select(self.all_projects[0])
        else:
            self.update_project_dropdown()
            self.update_status("Info", "No projects found. Create a new project.", "blue")

        self.on_system_change(vars_['system'].get()) # Initial call to set visibility

    # --- New Project Management Methods ---
    def load_project_list(self):
        """Loads the list of all project names."""
        if os.path.exists(MASTER_PROJECT_LIST_FILE):
            try:
                with open(MASTER_PROJECT_LIST_FILE, 'r') as f:
                    self.all_projects = json.load(f)
                self.update_project_dropdown()
            except Exception as e:
                self.update_status("Error", f"Could not load project list: {e}", "red")
        else:
            self.all_projects = []
            self.update_project_dropdown()

    def save_project_list(self):
        """Saves the current list of project names."""
        with open(MASTER_PROJECT_LIST_FILE, 'w') as f:
            json.dump(self.all_projects, f, indent=4)

    def update_project_dropdown(self):
        """Updates the project selection dropdown with current project names."""
        # Ensure the values are always a list of strings
        self.project_dropdown.configure(values=self.all_projects)
        if self.all_projects and self.current_project_name in self.all_projects:
            self.vars['selected_project'].set(self.current_project_name)
        elif self.all_projects:
            # If current_project_name is not in list (e.g., deleted), set to first
            self.vars['selected_project'].set(self.all_projects[0])
            self.current_project_name = self.all_projects[0]
            self.set_current_project_paths() # Ensure paths are set for the default selected project
        else:
            self.vars['selected_project'].set("") # No projects to select

    def create_new_project(self):
        """Creates a new project, setting up its files and making it the current project."""
        new_name = self.vars['new_project_name'].get().strip()
        if not new_name:
            self.update_status("Error", "Please enter a name for the new project.", "red")
            return
        
        if new_name in self.all_projects:
            self.update_status("Error", f"Project '{new_name}' already exists.", "red")
            return
        
        # Add new project to the list and save
        self.all_projects.append(new_name)
        self.save_project_list()
        self.update_project_dropdown()

        # Set this new project as current
        self.vars['selected_project'].set(new_name)
        self.on_project_select(new_name) # This will create the initial files

        self.update_status("Created", f"New project '{new_name}'", "green")
        self.vars['new_project_name'].set("") # Clear new project name field

    def on_project_select(self, project_name):
        """Handles selection of a project from the dropdown."""
        self.current_project_name = project_name
        self.set_current_project_paths()
        self.load_saved_elevations_for_current_project()
        self.clear_form() # Clear the form when switching projects
        self.update_status("Switched to", f"project '{project_name}'", "blue")

    def set_current_project_paths(self):
        """Defines the project-specific file paths and ensures initial files exist."""
        if not self.current_project_name:
            # If no project is selected, default to dummy paths or handle appropriately
            self.current_excel_path = os.path.join(PROJECTS_DIR, "default_report.xlsx")
            self.current_elevations_json_path = os.path.join(PROJECTS_DIR, "default_elevations.json")
            self.current_extra_materials_json_path = os.path.join(PROJECTS_DIR, "default_extra_materials.json")
            return

        # Sanitize project name for file paths
        base_name = os.path.join(PROJECTS_DIR, self.current_project_name.replace(" ", "_").replace("/", "_").replace("\\", "_"))
        self.current_excel_path = f"{base_name}_Report.xlsx"
        self.current_elevations_json_path = f"{base_name}_Elevations.json"
        self.current_extra_materials_json_path = f"{base_name}_ExtraMaterials.json"

        # Ensure that if a new project is created, its initial Excel and extra_materials files exist
        if not os.path.exists(self.current_excel_path):
            wb = Workbook()
            ws = wb.active
            ws.title = "Report" # Set a default sheet name
            wb.save(self.current_excel_path)
            print(f"Created initial Excel file for '{self.current_project_name}' at {self.current_excel_path}")
        
        if not os.path.exists(self.current_extra_materials_json_path):
            with open(self.current_extra_materials_json_path, 'w') as f:
                json.dump({}, f, indent=4) # Empty extra materials for a new project
            print(f"Created initial Extra Materials JSON for '{self.current_project_name}' at {self.current_extra_materials_json_path}")


    def load_saved_elevations_for_current_project(self):
        """Loads saved elevations for the currently selected project."""
        # Use the project-specific path
        if os.path.exists(self.current_elevations_json_path):
            try:
                with open(self.current_elevations_json_path, 'r') as f:
                    self.saved_elevations = json.load(f)
                self.update_saved_elevation_dropdown()
                self.update_status("Loaded", f"elevations for '{self.current_project_name}'", "green")
            except Exception as e:
                self.update_status("Error", f"Could not load elevations for '{self.current_project_name}': {e}", "red")
        else:
            self.saved_elevations = {}
            self.update_saved_elevation_dropdown()
            self.update_status("Info", f"No elevations found for '{self.current_project_name}'.", "blue")
            # Create an empty JSON file if it doesn't exist for the current project
            with open(self.current_elevations_json_path, 'w') as f:
                json.dump({}, f, indent=4)
            print(f"Created initial Elevations JSON for '{self.current_project_name}' at {self.current_elevations_json_path}")


    # --- Existing Methods Modified to Use Project Paths ---

    def on_system_change(self, selected):
        """Adjusts visibility of bay input fields based on system selection."""
        show_bays = selected == "YES 45TU FRONT SET(OG)"
        # Adjusted row numbers for bay fields due to new project section
        for field, row in [('bays_wide', 8), ('bays_tall', 9)]:
            # Ensure the widget keys match the updated ones in __init__
            label_widget = self.widgets.get(f"label_{row}")
            entry_widget = self.widgets.get(f"entry_{field}")
            
            if label_widget and entry_widget: # Check if widgets exist
                if show_bays:
                    label_widget.grid(row=row, column=0, sticky="w", pady=5)
                    entry_widget.grid(row=row, column=1, sticky="ew", pady=5)
                else:
                    label_widget.grid_forget()
                    entry_widget.grid_forget()
            else:
                print(f"Warning: Widget for {field} at row {row} not found in self.widgets.")


    def on_saved_elevation_select(self, elev_type):
        """Loads selected elevation data into the form fields."""
        if elev_type not in self.saved_elevations:
            return
        data = self.saved_elevations[elev_type]
        for key, var_key in [
            ('system', 'system'),
            ('finish', 'finish'),
            ('door_size', 'door'),
            ('total_count', 'total_count'),
            ('bays_wide', 'bays_wide'),
            ('bays_tall', 'bays_tall'),
            ('opening_width_inches', 'opening_width'),
            ('opening_height_inches', 'opening_height'),
        ]:
            # Ensure values are converted to string for StringVar.set()
            self.vars[var_key].set(str(data.get(key, '')))
        self.vars['elevation_type'].set(elev_type) # Ensure elevation_type entry is updated
        self.on_system_change(self.vars['system'].get()) # Update visibility based on loaded system
        self.update_status("Loaded", elev_type, "green")

    def save_elevation_type(self):
        """Saves or updates an elevation and generates the Excel report."""
        if not self.current_project_name:
            self.update_status("Error", "Please create or select a project first.", "red")
            return

        try:
            v = self.vars
            elev = v['elevation_type'].get().strip()
            if not elev:
                self.update_status("Error", "Please enter an elevation type.", "red")
                return
            system = v['system'].get()
            finish = v['finish'].get()
            door_size = v['door'].get()
            total = int(v['total_count'].get())
            ow = float(v['opening_width'].get())
            oh = float(v['opening_height'].get())
            bays_wide = int(v['bays_wide'].get()) if system == self.system_options[0] else 0
            bays_tall = int(v['bays_tall'].get()) if system == self.system_options[0] else 0

            calculated = []
            if system == self.system_options[0]:
                calculated = calculate_yes45tu_quantities(bays_wide, bays_tall, total, ow, oh, door_size)

            sqft_per = calculate_rectangle_area(ow / 12, oh / 12)
            total_sqft = sqft_per * total
            perimeter = calculate_perimeter(ow / 12, oh / 12)
            total_perimeter = perimeter * total

            # Pass the current project's file paths to generate_excel_report
            generate_excel_report(
                excel_path=self.current_excel_path,
                elevations_json_path=self.current_elevations_json_path,
                extra_materials_json_path=self.current_extra_materials_json_path,
                system_input=system,
                finish_input=finish,
                elevation_type=elev,
                total_count=total,
                bays_wide=bays_wide,
                bays_tall=bays_tall,
                opening_width=ow,
                opening_height=oh,
                sqft_per_type=sqft_per,
                total_sqft=total_sqft,
                perimeter_ft=perimeter,
                total_perimeter_ft=total_perimeter,
                calculated_outputs=calculated,
                completion_callback=lambda msg=None: self.update_status("Report", msg, "green"),
                door_size=door_size,
                mode="save_or_update" # Explicitly indicate save/update mode
            )
            
            # After generate_excel_report completes, reload saved_elevations to reflect changes
            # This ensures the dropdown is always up-to-date with the current project's elevations
            self.load_saved_elevations_for_current_project() 
            self.vars['saved_elevation_types'].set(elev) # Select the newly saved/updated elevation
            self.update_status("Saved", elev, "green")

        except ValueError as e:
            self.update_status("Error", str(e), "red")
        except Exception as e:
            self.update_status("Error", f"An unexpected error occurred: {e}", "red")


    def delete_elevation_type(self):
        """Deletes a selected elevation and updates the Excel report."""
        if not self.current_project_name:
            self.update_status("Error", "Please select a project first.", "red")
            return

        elev = self.vars['saved_elevation_types'].get()
        if elev: # Ensure an elevation is selected
            if elev in self.saved_elevations:
                # Pass the current project's file paths to generate_excel_report
                generate_excel_report(
                    excel_path=self.current_excel_path,
                    elevations_json_path=self.current_elevations_json_path,
                    extra_materials_json_path=self.current_extra_materials_json_path,
                    system_input="", # These inputs are ignored during deletion mode
                    finish_input="",
                    elevation_type="",
                    total_count=0,
                    bays_wide=0,
                    bays_tall=0,
                    opening_width=0.0,
                    opening_height=0.0,
                    sqft_per_type=0.0,
                    total_sqft=0.0,
                    perimeter_ft=0.0,
                    total_perimeter_ft=0.0,
                    calculated_outputs=[],
                    completion_callback=lambda msg=None: self.update_status("Report", msg, "green"),
                    delete_elevation_type=elev # This is the key parameter for deletion
                )
                
                # After generate_excel_report completes, reload saved_elevations to reflect changes
                self.load_saved_elevations_for_current_project() # This will also update the dropdown
                self.clear_form() # Clear the form after deletion
                self.update_status("Deleted", elev, "green")
            else:
                self.update_status("Error", f"Elevation '{elev}' not found to delete.", "red")
        else:
            self.update_status("Error", "No elevation selected to delete.", "red")

    def delete_current_project(self):
        """Deletes the currently selected project and all its associated files."""
        if not self.current_project_name:
            self.update_status("Error", "No project selected to delete.", "red")
            return

        # Confirmation dialog (optional but recommended for deletion)
        if not tk.messagebox.askyesno(
            "Confirm Deletion",
            f"Are you sure you want to delete project '{self.current_project_name}' and all its data (Excel, JSONs)? This cannot be undone."
        ):
            self.update_status("Cancelled", "Project deletion cancelled.", "blue")
            return

        try:
            # Delete project-specific files
            files_to_delete = [
                self.current_excel_path,
                self.current_elevations_json_path,
                self.current_extra_materials_json_path
            ]
            
            for file_path in files_to_delete:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
            
            # Remove project from master list
            if self.current_project_name in self.all_projects:
                self.all_projects.remove(self.current_project_name)
                self.save_project_list()
                self.update_project_dropdown() # Update the dropdown

            self.update_status("Deleted", f"Project '{self.current_project_name}' and its files.", "green")
            
            # After deletion, select the first available project or clear the form
            if self.all_projects:
                self.vars['selected_project'].set(self.all_projects[0])
                self.on_project_select(self.all_projects[0])
            else:
                self.vars['selected_project'].set("")
                self.current_project_name = ""
                self.clear_form()
                self.saved_elevations = {}
                self.update_saved_elevation_dropdown()
                self.update_status("Info", "No projects remaining. Create a new project.", "blue")

        except Exception as e:
            self.update_status("Error", f"Could not delete project '{self.current_project_name}': {e}", "red")

    def update_saved_elevation_dropdown(self):
        """Updates the dropdown menu with saved elevation types for the current project."""
        keys = sorted(self.saved_elevations.keys())
        self.saved_elevation_dropdown.configure(values=keys)
        current = self.vars['saved_elevation_types'].get()
        if current not in keys:
            self.vars['saved_elevation_types'].set(keys[0] if keys else "")

    def clear_form(self):
        """Clears all input fields in the form."""
        for var in ['elevation_type', 'total_count', 'bays_wide', 'bays_tall', 'opening_width', 'opening_height']:
            self.vars[var].set("")
        self.vars['system'].set(self.system_options[0])
        self.vars['finish'].set(self.finish_options[0])
        self.vars['door'].set(self.door_options[0])
        self.on_system_change(self.vars['system'].get())

    def update_status(self, action, message, color="red"):
        """Updates the status label with a given message and color."""
        full_message = f"{action}: {message}"
        self.status_label.configure(text=full_message, text_color=color)

if __name__ == "__main__":
    app = App()
    app.after(10, lambda: app.state('zoomed'))
    app.mainloop()