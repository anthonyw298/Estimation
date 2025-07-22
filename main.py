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

        # Variables for inputs
        self.var_system = tk.StringVar(value=self.system_options[0])
        self.var_finish = tk.StringVar(value=self.finish_options[0])
        self.var_elevation_type = tk.StringVar()
        self.var_total_count = tk.StringVar()
        self.var_bays_wide = tk.StringVar()
        self.var_bays_tall = tk.StringVar()
        self.var_opening_width = tk.StringVar()
        self.var_opening_height = tk.StringVar()

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

        # Elevation Type
        ctk.CTkLabel(self.main_frame, text="Elevation Type:").grid(row=2, column=0, sticky="w", pady=5)
        self.entry_elevation_type = ctk.CTkEntry(self.main_frame, textvariable=self.var_elevation_type)
        self.entry_elevation_type.grid(row=2, column=1, sticky="ew", pady=5)

        # Total Count
        ctk.CTkLabel(self.main_frame, text="Total Count:").grid(row=3, column=0, sticky="w", pady=5)
        self.entry_total_count = ctk.CTkEntry(self.main_frame, textvariable=self.var_total_count)
        self.entry_total_count.grid(row=3, column=1, sticky="ew", pady=5)

        # Bays Wide (only for YES 45TU)
        self.label_bays_wide = ctk.CTkLabel(self.main_frame, text="# Bays Wide:")
        self.entry_bays_wide = ctk.CTkEntry(self.main_frame, textvariable=self.var_bays_wide)

        # Bays Tall (only for YES 45TU)
        self.label_bays_tall = ctk.CTkLabel(self.main_frame, text="# Bays Tall:")
        self.entry_bays_tall = ctk.CTkEntry(self.main_frame, textvariable=self.var_bays_tall)

        # Opening Width
        ctk.CTkLabel(
            self.main_frame, text="Opening Width (in inches):"
        ).grid(row=6, column=0, sticky="w", pady=5)
        self.entry_opening_width = ctk.CTkEntry(self.main_frame, textvariable=self.var_opening_width)
        self.entry_opening_width.grid(row=6, column=1, sticky="ew", pady=5)

        # Opening Height
        ctk.CTkLabel(
            self.main_frame, text="Opening Height (in inches):"
        ).grid(row=7, column=0, sticky="w", pady=5)
        self.entry_opening_height = ctk.CTkEntry(self.main_frame, textvariable=self.var_opening_height)
        self.entry_opening_height.grid(row=7, column=1, sticky="ew", pady=5)

        # Buttons Frame
        self.buttons_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.buttons_frame.grid(row=8, column=0, columnspan=2, sticky="e", pady=20, padx=(0, 30))

        self.submit_button = ctk.CTkButton(
            self.buttons_frame,
            text="Generate New Excel Report",
            command=lambda: self.submit_data(mode="new")
        )
        self.submit_button.pack(side="left", padx=(0, 20))

        self.add_elevation_button = ctk.CTkButton(
            self.buttons_frame,
            text="Add New Elevation Type to Current Excel",
            command=lambda: self.submit_data(mode="append")
        )
        self.add_elevation_button.pack(side="left")

        # Status Label
        self.status_label = ctk.CTkLabel(self.main_frame, text="", text_color="red")
        self.status_label.grid(row=9, column=0, columnspan=2)

        # Configure grid weights for responsiveness
        self.main_frame.grid_columnconfigure(1, weight=1)

        # Initialize UI based on default system selection
        self.on_system_change(self.var_system.get())

    def on_system_change(self, selected_system):
        if selected_system == "YES 45TU FRONT SET(OG)":
            self.label_bays_wide.grid(row=4, column=0, sticky="w", pady=5)
            self.entry_bays_wide.grid(row=4, column=1, sticky="ew", pady=5)
            self.label_bays_tall.grid(row=5, column=0, sticky="w", pady=5)
            self.entry_bays_tall.grid(row=5, column=1, sticky="ew", pady=5)
        else:
            self.label_bays_wide.grid_forget()
            self.entry_bays_wide.grid_forget()
            self.label_bays_tall.grid_forget()
            self.entry_bays_tall.grid_forget()

    def submit_data(self, mode="new"):
        try:
            system_input = self.var_system.get()
            finish_input = self.var_finish.get()
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
                finish_input,
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
                completion_callback=self.update_status,
                mode=mode
            )
            self.update_status(
                "Excel report generated successfully!" if mode == "new" else "Elevation type added successfully!",
                "green"
            )

        except ValueError as e:
            self.update_status(f"Input error: {e}. Please enter valid numbers.", "red")

    def update_status(self, message, color="red"):
        self.status_label.configure(text=message, text_color=color)


if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    app = App()
    app.after(10, lambda: app.state('zoomed'))
    app.mainloop()
