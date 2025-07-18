import customtkinter as ctk
import tkinter as tk
from utils.excel_generator import generate_excel_report
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Data Entry Application")
        self.state('zoomed')  # fullscreen maximized
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.system_options = ["YES 45TU FRONT SET(OG)", "Other"]
        self.finish_options = ["Clear", "Black", "Painted"]

        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=0, padx=40, pady=40, sticky="nsew")
        self.main_frame.grid_columnconfigure(1, weight=1)

        # System Dropdown
        ctk.CTkLabel(self.main_frame, text="Select System Input:", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", pady=(0,10))
        self.system_var = tk.StringVar(value=self.system_options[0])
        self.system_dropdown = ctk.CTkComboBox(self.main_frame, values=self.system_options, variable=self.system_var, width=300, command=self.system_changed)
        self.system_dropdown.grid(row=0, column=1, sticky="ew", pady=(0,10))

        # Finish Dropdown
        ctk.CTkLabel(self.main_frame, text="Select Finish:", font=ctk.CTkFont(size=16, weight="bold")).grid(row=1, column=0, sticky="w", pady=(0,10))
        self.finish_var = tk.StringVar(value=self.finish_options[0])
        self.finish_dropdown = ctk.CTkComboBox(self.main_frame, values=self.finish_options, variable=self.finish_var, width=300)
        self.finish_dropdown.grid(row=1, column=1, sticky="ew", pady=(0,10))

        # Input fields dictionary
        self.inputs = {}

        # Labels order with bays_wide and bays_tall moved up above opening height and width
        labels_common = [
            ("Elevation Type:", "elevation_type"),
            ("Total Count:", "total_count"),
            ("# Bays Wide:", "bays_wide"),
            ("# Bays Tall:", "bays_tall"),
            ("Opening Width (in inches):", "opening_width"),
            ("Opening Height (in inches):", "opening_height"),
        ]

        row_idx = 2
        for label_text, key in labels_common:
            ctk.CTkLabel(self.main_frame, text=label_text).grid(row=row_idx, column=0, sticky="w", pady=5)
            entry = ctk.CTkEntry(self.main_frame, width=300)
            entry.grid(row=row_idx, column=1, sticky="ew", pady=5)
            self.inputs[key] = entry
            row_idx += 1

        # Submit Button
        self.submit_button = ctk.CTkButton(self.main_frame, text="Submit", command=self.submit_data)
        self.submit_button.grid(row=row_idx, column=0, columnspan=2, pady=(20, 10))

        # Status label
        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="red")
        self.status_label.grid(row=row_idx+1, column=0, columnspan=2, pady=(0, 10))

        self.system_changed(self.system_var.get())

    def system_changed(self, system_name):
        # Show/hide bays inputs only if YES 45TU selected
        if system_name == "YES 45TU FRONT SET(OG)":
            self.inputs["bays_wide"].master.grid()
            self.inputs["bays_tall"].master.grid()
        else:
            self.inputs["bays_wide"].master.grid_remove()
            self.inputs["bays_tall"].master.grid_remove()

        # Clear inputs and status
        self.status_label.configure(text="")
        for entry in self.inputs.values():
            entry.delete(0, tk.END)

    def submit_data(self):
        system_input = self.system_var.get()
        finish = self.finish_var.get()

        try:
            elevation_type = self.inputs["elevation_type"].get().strip()
            if not elevation_type:
                raise ValueError("Elevation Type cannot be empty.")

            total_count = int(self.inputs["total_count"].get().strip())

            bays_wide = 0
            bays_tall = 0
            if system_input == "YES 45TU FRONT SET(OG)":
                bays_wide = int(self.inputs["bays_wide"].get().strip())
                bays_tall = int(self.inputs["bays_tall"].get().strip())

            opening_width_inches = float(self.inputs["opening_width"].get().strip())
            opening_height_inches = float(self.inputs["opening_height"].get().strip())

            calculated_outputs = []
            if system_input == "YES 45TU FRONT SET(OG)":
                calculated_outputs = calculate_yes45tu_quantities(
                    bays_wide, bays_tall, total_count, opening_width_inches, opening_height_inches
                )

            opening_width_feet = opening_width_inches / 12.0
            opening_height_feet = opening_height_inches / 12.0

            sqft_per_type = calculate_rectangle_area(opening_width_feet, opening_height_feet)
            total_sqft = sqft_per_type * total_count
            perimeter_ft = calculate_perimeter(opening_width_feet, opening_height_feet)
            total_perimeter_ft = perimeter_ft * total_count

        except ValueError as e:
            self.status_label.configure(text=f"Input Error: {e}")
            return

        def update_status(msg, color):
            self.status_label.configure(text=msg, text_color=color)

        generate_excel_report(
            system_input,
            elevation_type,
            total_count,
            bays_wide,
            bays_tall,
            opening_width_inches,
            opening_height_inches,
            sqft_per_type,
            total_sqft,
            perimeter_ft,
            total_perimeter_ft,
            calculated_outputs,
            update_status,
            finish
        )


if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    app = App()
    app.mainloop()
