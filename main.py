import customtkinter as ctk
import tkinter as tk
from utils.excel_generator import generate_excel_report
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("United Glass Estimation Calculation Tool")
        self.state('zoomed')  # Maximize window

        # SYSTEM OPTIONS
        self.system_options = ["YES 45TU FRONT SET(OG)", "Other"]
        self.finish_options = ["Clear", "Black", "Paint"]

        # Saved elevations dictionary: elevation_type -> data dict
        self.saved_elevations = {}

        # Variables for inputs
        self.var_system = tk.StringVar(value=self.system_options[0])
        self.var_finish = tk.StringVar(value=self.finish_options[0])
        self.var_elevation_type = tk.StringVar()
        self.var_total_count = tk.StringVar()
        self.var_bays_wide = tk.StringVar()
        self.var_bays_tall = tk.StringVar()
        self.var_opening_width = tk.StringVar()
        self.var_opening_height = tk.StringVar()

        # Variable for saved elevation types dropdown
        self.var_saved_elevation_types = tk.StringVar()

        # Main Frame
        self.main_frame = ctk.CTkFrame(self, corner_radius=20)
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=30)

        # System Dropdown
        ctk.CTkLabel(
            self.main_frame, text="Select System:", font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, sticky="w", pady=(0, 15))
        self.system_dropdown = ctk.CTkOptionMenu(
            self.main_frame,
            values=self.system_options,
            variable=self.var_system,
            command=self.on_system_change
        )
        self.system_dropdown.grid(row=0, column=1, sticky="ew", pady=(0, 15))

        # Finish Dropdown
        ctk.CTkLabel(
            self.main_frame, text="Select Finish:", font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=1, column=0, sticky="w", pady=(0, 15))
        self.finish_dropdown = ctk.CTkOptionMenu(
            self.main_frame, values=self.finish_options, variable=self.var_finish
        )
        self.finish_dropdown.grid(row=1, column=1, sticky="ew", pady=(0, 15))

        # Saved Elevation Types Dropdown (for selecting/editing)
        ctk.CTkLabel(
            self.main_frame, text="Saved Elevation Types:", font=ctk.CTkFont(size=14)
        ).grid(row=2, column=0, sticky="w", pady=(5, 15))
        self.saved_elevation_dropdown = ctk.CTkOptionMenu(
            self.main_frame,
            values=[],
            variable=self.var_saved_elevation_types,
            command=self.on_saved_elevation_select
        )
        self.saved_elevation_dropdown.grid(row=2, column=1, sticky="ew", pady=(5, 15))

        # Elevation Type Entry (for new or editing)
        ctk.CTkLabel(self.main_frame, text="Elevation Type:").grid(row=3, column=0, sticky="w", pady=5)
        self.entry_elevation_type = ctk.CTkEntry(self.main_frame, textvariable=self.var_elevation_type)
        self.entry_elevation_type.grid(row=3, column=1, sticky="ew", pady=5)

        # Total Count
        ctk.CTkLabel(self.main_frame, text="Total Count:").grid(row=4, column=0, sticky="w", pady=5)
        self.entry_total_count = ctk.CTkEntry(self.main_frame, textvariable=self.var_total_count)
        self.entry_total_count.grid(row=4, column=1, sticky="ew", pady=5)

        # Bays Wide (only for YES 45TU)
        self.label_bays_wide = ctk.CTkLabel(self.main_frame, text="# Bays Wide:")
        self.entry_bays_wide = ctk.CTkEntry(self.main_frame, textvariable=self.var_bays_wide)

        # Bays Tall (only for YES 45TU)
        self.label_bays_tall = ctk.CTkLabel(self.main_frame, text="# Bays Tall:")
        self.entry_bays_tall = ctk.CTkEntry(self.main_frame, textvariable=self.var_bays_tall)

        # Opening Width
        ctk.CTkLabel(
            self.main_frame, text="Opening Width (in inches):"
        ).grid(row=7, column=0, sticky="w", pady=5)
        self.entry_opening_width = ctk.CTkEntry(self.main_frame, textvariable=self.var_opening_width)
        self.entry_opening_width.grid(row=7, column=1, sticky="ew", pady=5)

        # Opening Height
        ctk.CTkLabel(
            self.main_frame, text="Opening Height (in inches):"
        ).grid(row=8, column=0, sticky="w", pady=5)
        self.entry_opening_height = ctk.CTkEntry(self.main_frame, textvariable=self.var_opening_height)
        self.entry_opening_height.grid(row=8, column=1, sticky="ew", pady=5)

        # Buttons Frame
        self.buttons_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.buttons_frame.grid(row=9, column=0, columnspan=2, sticky="e", pady=20, padx=(0, 30))

        self.submit_button = ctk.CTkButton(
            self.buttons_frame,
            text="Save Elevation Type",
            command=self.save_elevation_type
        )
        self.submit_button.pack(side="left", padx=(0, 20))

        self.delete_button = ctk.CTkButton(
            self.buttons_frame,
            text="Delete Elevation",
            command=self.delete_elevation_type
        )
        self.delete_button.pack(side="left", padx=(0, 20))

        # Status Label
        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="red")
        self.status_label.grid(row=10, column=0, columnspan=2)

        # Configure grid weights for responsiveness
        self.main_frame.grid_columnconfigure(1, weight=1)

        # Initialize UI based on default system selection
        self.on_system_change(self.var_system.get())

    def on_system_change(self, selected_system):
        if selected_system == "YES 45TU FRONT SET(OG)":
            self.label_bays_wide.grid(row=5, column=0, sticky="w", pady=5)
            self.entry_bays_wide.grid(row=5, column=1, sticky="ew", pady=5)
            self.label_bays_tall.grid(row=6, column=0, sticky="w", pady=5)
            self.entry_bays_tall.grid(row=6, column=1, sticky="ew", pady=5)
        else:
            self.label_bays_wide.grid_forget()
            self.entry_bays_wide.grid_forget()
            self.label_bays_tall.grid_forget()
            self.entry_bays_tall.grid_forget()

    def on_saved_elevation_select(self, elevation_type):
        if not elevation_type or elevation_type not in self.saved_elevations:
            return
        data = self.saved_elevations[elevation_type]

        self.var_elevation_type.set(elevation_type)
        self.var_system.set(data.get("system", self.system_options[0]))
        self.var_finish.set(data.get("finish", self.finish_options[0]))
        self.var_total_count.set(str(data.get("total_count", "")))
        self.var_bays_wide.set(str(data.get("bays_wide", "")))
        self.var_bays_tall.set(str(data.get("bays_tall", "")))
        self.var_opening_width.set(str(data.get("opening_width_inches", "")))
        self.var_opening_height.set(str(data.get("opening_height_inches", "")))

        self.on_system_change(self.var_system.get())
        self.update_status(f"Loaded elevation '{elevation_type}' for editing.", "green")

    def save_elevation_type(self):
        try:
            elevation_type = self.var_elevation_type.get().strip()
            if not elevation_type:
                self.update_status("Please enter an elevation type.", "red")
                return

            system_input = self.var_system.get()
            finish_input = self.var_finish.get()
            total_count = int(self.var_total_count.get())
            opening_width_inches = float(self.var_opening_width.get())
            opening_height_inches = float(self.var_opening_height.get())

            bays_wide = 0
            bays_tall = 0
            calculated_outputs = []

            if system_input == "YES 45TU FRONT SET(OG)":
                bays_wide = int(self.var_bays_wide.get())
                bays_tall = int(self.var_bays_tall.get())
                calculated_outputs = calculate_yes45tu_quantities(
                    bays_wide, bays_tall, total_count, opening_width_inches, opening_height_inches
                )

            opening_width_feet = opening_width_inches / 12.0
            opening_height_feet = opening_height_inches / 12.0

            sqft_per_type = calculate_rectangle_area(opening_width_feet, opening_height_feet)
            total_sqft = sqft_per_type * total_count
            perimeter_ft = calculate_perimeter(opening_width_feet, opening_height_feet)
            total_perimeter_ft = perimeter_ft * total_count

            self.saved_elevations[elevation_type] = {
                "system": system_input,
                "finish": finish_input,
                "total_count": total_count,
                "bays_wide": bays_wide,
                "bays_tall": bays_tall,
                "opening_width_inches": opening_width_inches,
                "opening_height_inches": opening_height_inches,
                "sqft_per_type": sqft_per_type,
                "total_sqft": total_sqft,
                "perimeter_ft": perimeter_ft,
                "total_perimeter_ft": total_perimeter_ft,
                "calculated_outputs": calculated_outputs,
            }

            self.update_saved_elevation_dropdown()
            self.var_saved_elevation_types.set(elevation_type)  # Select saved elevation
            self.update_status(f"Saved elevation '{elevation_type}' successfully.", "green")

            # Regenerate Excel report with all saved elevations
            all_elevations_data = list(self.saved_elevations.values())
            generate_excel_report(
                system_input=system_input,
                finish_input=finish_input,
                elevation_type=elevation_type,
                total_count=total_count,
                bays_wide=bays_wide,
                bays_tall=bays_tall,
                opening_width=opening_width_inches,
                opening_height=opening_height_inches,
                sqft_per_type=sqft_per_type,
                total_sqft=total_sqft,
                perimeter_ft=perimeter_ft,
                total_perimeter_ft=total_perimeter_ft,
                calculated_outputs=calculated_outputs,
                all_elevations=all_elevations_data,
                completion_callback=self.update_status,
                mode="regenerate"
            )

        except ValueError as e:
            self.update_status(f"Invalid input: {e}", "red")

    def delete_elevation_type(self):
        elevation_type = self.var_saved_elevation_types.get()
        if elevation_type and elevation_type in self.saved_elevations:
            del self.saved_elevations[elevation_type]
            self.update_saved_elevation_dropdown()
            self.clear_form()
            self.update_status(f"Deleted elevation '{elevation_type}'.", "green")

            # Pass all elevations except deleted to regenerate the Excel report
            all_elevations_data = list(self.saved_elevations.values())
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
                all_elevations=all_elevations_data,
                completion_callback=self.update_status,
                mode="regenerate",
                delete_elevation_type=elevation_type  # Tell Excel generator which was deleted
            )
        else:
            self.update_status("No elevation selected to delete.", "red")

    def update_saved_elevation_dropdown(self):
        elevation_types = sorted(self.saved_elevations.keys())
        self.saved_elevation_dropdown.configure(values=elevation_types)
        if elevation_types:
            # Keep selection if possible
            current = self.var_saved_elevation_types.get()
            if current not in elevation_types:
                self.var_saved_elevation_types.set(elevation_types[0])
        else:
            self.var_saved_elevation_types.set("")

    def clear_form(self):
        self.var_elevation_type.set("")
        self.var_total_count.set("")
        self.var_bays_wide.set("")
        self.var_bays_tall.set("")
        self.var_opening_width.set("")
        self.var_opening_height.set("")
        self.var_system.set(self.system_options[0])
        self.var_finish.set(self.finish_options[0])
        self.on_system_change(self.var_system.get())

    def update_status(self, message=' Sucessfully Deleted Elevation Type', color='white'):
        self.status_label.configure(text=message, text_color=color)


if __name__ == "__main__":
    app = App()
    app.mainloop()
