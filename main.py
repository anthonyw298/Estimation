import customtkinter as ctk
import tkinter as tk
import json
import os

from utils.excel_generator import generate_excel_report
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter

SAVE_FILE = "saved_elevations.json"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("United Glass Estimation Calculation Tool")
        self.state('zoomed')

        # Constants
        self.system_options = ["YES 45TU FRONT SET(OG)", "Other"]
        self.finish_options = ["Clear", "Black", "Paint"]
        self.door_options = ["3' X 7'", "3' X 8'", "3' X 9'", "6' X 7'", "6' X 8'", "6' X 9'"]
        self.saved_elevations = {}

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
        )
        self.vars = vars_

        # UI setup
        self.main_frame = ctk.CTkFrame(self, corner_radius=20)
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=30)
        self.main_frame.grid_columnconfigure(1, weight=1)

        labels = [
            ("Select System:", 0),
            ("Select Finish:", 1),
            ("Select Door Size:", 2),
            ("Saved Elevation Types:", 3),
            ("Elevation Type:", 4),
            ("Total Count:", 5),
            ("# Bays Wide:", 6),
            ("# Bays Tall:", 7),
            ("Opening Width (in inches):", 8),
            ("Opening Height (in inches):", 9),
        ]
        self.widgets = {}
        for text, row in labels:
            lbl = ctk.CTkLabel(self.main_frame, text=text)
            lbl.grid(row=row, column=0, sticky="w", pady=5)
            self.widgets[f"label_{row}"] = lbl

        self.system_dropdown = ctk.CTkOptionMenu(
            self.main_frame,
            values=self.system_options,
            variable=vars_['system'],
            command=self.on_system_change
        )
        self.system_dropdown.grid(row=0, column=1, sticky="ew", pady=5)

        self.finish_dropdown = ctk.CTkOptionMenu(
            self.main_frame, values=self.finish_options, variable=vars_['finish']
        )
        self.finish_dropdown.grid(row=1, column=1, sticky="ew", pady=5)

        self.door_dropdown = ctk.CTkOptionMenu(
            self.main_frame, values=self.door_options, variable=vars_['door']
        )
        self.door_dropdown.grid(row=2, column=1, sticky="ew", pady=5)

        self.saved_elevation_dropdown = ctk.CTkOptionMenu(
            self.main_frame,
            values=[],
            variable=vars_['saved_elevation_types'],
            command=self.on_saved_elevation_select
        )
        self.saved_elevation_dropdown.grid(row=3, column=1, sticky="ew", pady=5)

        entry_fields = ['elevation_type', 'total_count', 'bays_wide', 'bays_tall', 'opening_width', 'opening_height']
        for idx, field in enumerate(entry_fields, start=4):
            entry = ctk.CTkEntry(self.main_frame, textvariable=vars_[field])
            entry.grid(row=idx, column=1, sticky="ew", pady=5)
            self.widgets[f"entry_{field}"] = entry

        btn_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        btn_frame.grid(row=10, column=0, columnspan=2, sticky="e", pady=20, padx=(0, 30))
        self.submit_button = ctk.CTkButton(btn_frame, text="Save Elevation Type", command=self.save_elevation_type)
        self.submit_button.pack(side="left", padx=(0, 20))
        self.delete_button = ctk.CTkButton(btn_frame, text="Delete Elevation", command=self.delete_elevation_type)
        self.delete_button.pack(side="left", padx=(0, 20))

        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="red")
        self.status_label.grid(row=11, column=0, columnspan=2)

        self.load_saved_elevations()
        self.on_system_change(vars_['system'].get())

    def on_system_change(self, selected):
        show_bays = selected == "YES 45TU FRONT SET(OG)"
        for field, row in [('bays_wide', 6), ('bays_tall', 7)]:
            widget = self.widgets[f"label_{row}"]
            entry = self.widgets[f"entry_{field}"]
            if show_bays:
                widget.grid(row=row, column=0, sticky="w", pady=5)
                entry.grid(row=row, column=1, sticky="ew", pady=5)
            else:
                widget.grid_forget()
                entry.grid_forget()

    def on_saved_elevation_select(self, elev_type):
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
            self.vars[var_key].set(str(data.get(key, '')))
        self.vars['elevation_type'].set(elev_type)
        self.on_system_change(self.vars['system'].get())
        self.update_status("Loaded", elev_type, "green")

    def save_elevation_type(self):
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

            self.saved_elevations[elev] = {
                "system": system,
                "finish": finish,
                "door_size": door_size,
                "total_count": total,
                "bays_wide": bays_wide,
                "bays_tall": bays_tall,
                "opening_width_inches": ow,
                "opening_height_inches": oh,
                "sqft_per_type": sqft_per,
                "total_sqft": total_sqft,
                "perimeter_ft": perimeter,
                "total_perimeter_ft": total_perimeter,
                "calculated_outputs": calculated,
            }

            self.update_saved_elevation_dropdown()
            self.vars['saved_elevation_types'].set(elev)
            self.save_elevations_to_disk()

            generate_excel_report(
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
                all_elevations=list(self.saved_elevations.values()),
                completion_callback=lambda msg=None: self.update_status("Report", msg, "green"),
                mode="regenerate",
            )
            self.update_status("Saved", elev, "green")
        except ValueError as e:
            self.update_status("Error", str(e), "red")

    def delete_elevation_type(self):
        elev = self.vars['saved_elevation_types'].get()
        if elev in self.saved_elevations:
            self.saved_elevations.pop(elev)
            self.update_saved_elevation_dropdown()
            self.clear_form()
            self.update_status("Deleted", elev, "green")
            self.save_elevations_to_disk()

            generate_excel_report(
                system_input="",
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
                all_elevations=list(self.saved_elevations.values()),
                completion_callback=lambda msg=None: self.update_status("Report", msg, "green"),
                mode="regenerate",
                delete_elevation_type=elev
            )
        else:
            self.update_status("Error", "No elevation selected to delete.", "red")

    def update_saved_elevation_dropdown(self):
        keys = sorted(self.saved_elevations.keys())
        self.saved_elevation_dropdown.configure(values=keys)
        current = self.vars['saved_elevation_types'].get()
        if current not in keys:
            self.vars['saved_elevation_types'].set(keys[0] if keys else "")

    def clear_form(self):
        for var in ['elevation_type', 'total_count', 'bays_wide', 'bays_tall', 'opening_width', 'opening_height']:
            self.vars[var].set("")
        self.vars['system'].set(self.system_options[0])
        self.vars['finish'].set(self.finish_options[0])
        self.vars['door'].set(self.door_options[0])
        self.on_system_change(self.vars['system'].get())

    def update_status(self, action, elevation_name, color="red"):
        message = f"{action} '{elevation_name}'"
        self.status_label.configure(text=message, text_color=color)

    def save_elevations_to_disk(self):
        with open(SAVE_FILE, 'w') as f:
            json.dump(self.saved_elevations, f, indent=4)

    def load_saved_elevations(self):
        if os.path.exists(SAVE_FILE):
            try:
                with open(SAVE_FILE, 'r') as f:
                    self.saved_elevations = json.load(f)
                self.update_saved_elevation_dropdown()
                self.update_status("Loaded", "saved elevations", "green")
            except Exception as e:
                self.update_status("Error", f"Could not load saved elevations: {e}", "red")


if __name__ == "__main__":
    app = App()
    app.after(10, lambda: app.state('zoomed'))
    app.mainloop()
