import customtkinter as ctk
import tkinter as tk
from utils.excel_generator import generate_excel_report
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Data Entry Application")
        self.state('zoomed')  # Maximize window
        self.minsize(600, 400)

        # SYSTEM OPTIONS
        self.system_options = ["YES 45TU FRONT SET(OG)", "Other"]

        # Variables for inputs
        self.var_system = tk.StringVar(value=self.system_options[0])
        self.var_elevation_type = tk.StringVar()
        self.var_total_count = tk.StringVar()
        self.var_bays_wide = tk.StringVar()
        self.var_bays_tall = tk.StringVar()
        self.var_opening_width = tk.StringVar()
        self.var_opening_height = tk.StringVar()

        # Main Frame
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.pack(fill="both", expand=True, padx=30, pady=30)

        # System Dropdown
        ctk.CTkLabel(self.main_frame, text="Select System:", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w", pady=(0, 15))
        self.system_dropdown = ctk.CTkOptionMenu(self.main_frame, values=self.system_options, variable=self.var_system, command=self.on_system_change)
        self.system_dropdown.grid(row=0, column=1, sticky="ew", pady=(0, 15))

        # Elevation Type
        ctk.CTkLabel(self.main_frame, text="Elevation Type:").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_elevation_type = ctk.CTkEntry(self.main_frame, textvariable=self.var_elevation_type)
        self.entry_elevation_type.grid(row=1, column=1, sticky="ew", pady=5)

        # Total Count
        ctk.CTkLabel(self.main_frame, text="Total Count:").grid(row=2, column=0, sticky="w", pady=5)
        self.entry_total_count = ctk.CTkEntry(self.main_frame, textvariable=self.var_total_count)
        self.entry_total_count.grid(row=2, column=1, sticky="ew", pady=5)

        # Bays Wide (only for YES 45TU)
        self.label_bays_wide = ctk.CTkLabel(self.main_frame, text="# Bays Wide:")
        self.entry_bays_wide = ctk.CTkEntry(self.main_frame, textvariable=self.var_bays_wide)

        # Bays Tall (only for YES 45TU)
        self.label_bays_tall = ctk.CTkLabel(self.main_frame, text="# Bays Tall:")
        self.entry_bays_tall = ctk.CTkEntry(self.main_frame, textvariable=self.var_bays_tall)

        # Opening Width
        ctk.CTkLabel(self.main_frame, text="Opening Width (in inches):").grid(row=5, column=0, sticky="w", pady=5)
        self.entry_opening_width = ctk.CTkEntry(self.main_frame, textvariable=self.var_opening_width)
        self.entry_opening_width.grid(row=5, column=1, sticky="ew", pady=5)

        # Opening Height
        ctk.CTkLabel(self.main_frame, text="Opening Height (in inches):").grid(row=6, column=0, sticky="w", pady=5)
        self.entry_opening_height = ctk.CTkEntry(self.main_frame, textvariable=self.var_opening_height)
        self.entry_opening_height.grid(row=6, column=1, sticky="ew", pady=5)

        # Submit button
        self.submit_button = ctk.CTkButton(self.main_frame, text="Generate Excel Report", command=self.submit_data)
        self.submit_button.grid(row=7, column=0, columnspan=2, pady=20)

        # Status label
        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="red")
        self.status_label.grid(row=8, column=0, columnspan=2)

        # Configure grid weights for responsiveness
        self.main_frame.grid_columnconfigure(1, weight=1)

        # Initialize UI based on default system selection
        self.on_system_change(self.var_system.get())

    def on_system_change(self, selected_system):
        # Show/hide bays wide and bays tall fields based on system
        if selected_system == "YES 45TU FRONT SET(OG)":
            self.label_bays_wide.grid(row=3, column=0, sticky="w", pady=5)
            self.entry_bays_wide.grid(row=3, column=1, sticky="ew", pady=5)
            self.label_bays_tall.grid(row=4, column=0, sticky="w", pady=5)
            self.entry_bays_tall.grid(row=4, column=1, sticky="ew", pady=5)
        else:
            self.label_bays_wide.grid_forget()
            self.entry_bays_wide.grid_forget()
            self.label_bays_tall.grid_forget()
            self.entry_bays_tall.grid_forget()

    def submit_data(self):
        try:
            system_input = self.var_system.get()
            elevation_type = self.var_elevation_type.get().strip()
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
                self.update_status
            )
            self.update_status("Excel report generated successfully!", "green")

        except ValueError as e:
            self.update_status(f"Input error: {e}. Please enter valid numbers.", "red")

    def update_status(self, message, color="red"):
        self.status_label.configure(text=message, text_color=color)


if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop()
