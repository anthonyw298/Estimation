# -----------------------------
# IMPORTS
# -----------------------------
import customtkinter as ctk
import tkinter as tk
from utils.excel_generator import generate_excel_report  # EXCEL LOGIC: External report generator
from systems.yes45tu_front_set import calculate_yes45tu_quantities  # EXCEL LOGIC: System-specific calculator


# -----------------------------
# CLASS: Main Application
# -----------------------------
class App(ctk.CTk):
    """
    A CustomTkinter application for collecting data points for Excel calculations
    one at a time, then performing calculations and generating an Excel file.
    """

    def __init__(self):
        super().__init__()

        # -----------------------------
        # USER INTERFACE: Window Config
        # -----------------------------
        self.title("Excel Data Entry Application")
        self.geometry("500x350")  # Larger window for prompts
        self.resizable(False, False)

        # Center main frame
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # -----------------------------
        # DATA MANAGEMENT: Prompts & Storage
        # -----------------------------
        self.data_prompts = [
            "Enter System Input (YES 45TU Front Set(OG) or other):",
            "Enter Elevation Type:",
            "Enter Total Count:",
            "Enter # Bays Wide:",
            "Enter # Bays Tall:",
            "Enter Opening Width:",
            "Enter Opening Height:",
            "Enter Sq Ft per Type:",
            "Enter Total Sq Ft:",
            "Enter Perimeter Ft:",
            "Enter Total Perimeter Ft:"
        ]
        self.current_prompt_index = 0
        self.collected_data = []  # User input storage

        # -----------------------------
        # USER INTERFACE: Widgets
        # -----------------------------
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)

        # Title Label
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white",
            wraplength=400
        )
        self.title_label.grid(row=0, column=0, pady=(20, 10), sticky="s")

        # Entry Box
        self.entry_box = ctk.CTkEntry(
            self.main_frame,
            width=300,
            placeholder_text="Type here..."
        )
        self.entry_box.grid(row=1, column=0, pady=10, sticky="n")

        # Submit Button
        self.submit_button = ctk.CTkButton(
            self.main_frame,
            text="Submit",
            command=self.submit_data
        )
        self.submit_button.grid(row=2, column=0, pady=(10, 20), sticky="n")

        # Status Label
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="red"
        )
        self.status_label.grid(row=3, column=0, pady=(0, 10))

        # -----------------------------
        # USER INTERFACE: Initial Prompt
        # -----------------------------
        self.display_current_prompt()


    # -----------------------------
    # USER INTERFACE: Display Prompt
    # -----------------------------
    def display_current_prompt(self):
        if self.current_prompt_index < len(self.data_prompts):
            self.title_label.configure(text=self.data_prompts[self.current_prompt_index])
            self.entry_box.delete(0, tk.END)
            self.entry_box.focus_set()
            self.status_label.configure(text="")
        else:
            self.show_completion_screen()

    # -----------------------------
    # DATA MANAGEMENT: Handle Submit
    # -----------------------------
    def submit_data(self):
        entered_text = self.entry_box.get().strip()

        if entered_text:
            self.collected_data.append(entered_text)
            self.current_prompt_index += 1
            self.display_current_prompt()
        else:
            self.status_label.configure(text="Please enter some text!")
            self.after(1500, lambda: self.status_label.configure(text=""))


    # -----------------------------
    # USER INTERFACE: Completion Screen
    # -----------------------------
    def show_completion_screen(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=0)
        self.main_frame.grid_rowconfigure(2, weight=0)
        self.main_frame.grid_rowconfigure(3, weight=0)

        self.completion_label = ctk.CTkLabel(
            self.main_frame,
            text="Processing data and generating Excel file...",
            font=ctk.CTkFont(size=18, weight="bold"),
            wraplength=300
        )
        self.completion_label.grid(row=0, column=0, pady=30, sticky="nsew")

        self.after(100, lambda: self.call_excel_generator())

    def update_completion_status(self, message: str, color: str):
        self.completion_label.configure(text=message, text_color=color)

    # -----------------------------
    # EXCEL LOGIC: Call Report Generator
    # -----------------------------
    def call_excel_generator(self):
        if len(self.collected_data) != len(self.data_prompts):
            self.update_completion_status("Error: Incorrect number of inputs provided.", "red")
            return

        try:
            system_input = self.collected_data[0]
            elevation_type = self.collected_data[1]
            total_count = int(self.collected_data[2])
            bays_wide = int(self.collected_data[3])
            bays_tall = int(self.collected_data[4])
            opening_width = float(self.collected_data[5])
            opening_height = float(self.collected_data[6])
            sqft_per_type = float(self.collected_data[7])
            total_sqft = float(self.collected_data[8])
            perimeter_ft = float(self.collected_data[9])
            total_perimeter_ft = float(self.collected_data[10])
        except ValueError as e:
            self.update_completion_status(f"Error: Invalid number format. ({e})", "red")
            return

        if system_input == "YES 45TU Front Set(OG)":
            calculated_outputs = calculate_yes45tu_quantities(
                bays_wide, bays_tall, total_count, opening_width, opening_height
            )
        else:
            calculated_outputs = {}
            self.update_completion_status(
                f"Warning: No calculations defined for '{system_input}'.",
                "orange"
            )

        generate_excel_report(
            system_input,
            elevation_type,
            total_count,
            bays_wide,
            bays_tall,
            opening_width,
            opening_height,
            sqft_per_type,
            total_sqft,
            perimeter_ft,
            total_perimeter_ft,
            calculated_outputs,
            self.update_completion_status
        )


# -----------------------------
# APPLICATION ENTRY POINT
# -----------------------------
if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    app = App()
    app.mainloop()
