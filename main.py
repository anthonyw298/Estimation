import customtkinter as ctk
import tkinter as tk
from utils.excel_generator import generate_excel_report
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Data Entry Application")
        self.geometry("500x350")
        self.resizable(False, False)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.initial_prompt = "Enter System Input (YES 45TU Front Set(OG) or other):"
        
        self.yes45tu_prompts = [
            "Enter Elevation Type:",
            "Enter Total Count:",
            "Enter # Bays Wide:",
            "Enter # Bays Tall:",
            "Enter Opening Width (in inches):",
            "Enter Opening Height (in inches):"
        ]
        
        self.generic_prompts = [
            "Enter Elevation Type:",
            "Enter Total Count:",
            "Enter Opening Width (in inches):",
            "Enter Opening Height (in inches):"
        ]

        self.data_prompts = [self.initial_prompt]
        self.current_prompt_index = 0
        self.collected_data = [] 
        # Ensure system_input_chosen is always initialized
        self.system_input_chosen = "" 

        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)

        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white",
            wraplength=400
        )
        self.title_label.grid(row=0, column=0, pady=(20, 10), sticky="s")

        self.entry_box = ctk.CTkEntry(
            self.main_frame,
            width=300,
            placeholder_text="Type here..."
        )
        self.entry_box.grid(row=1, column=0, pady=10, sticky="n")

        self.submit_button = ctk.CTkButton(
            self.main_frame,
            text="Submit",
            command=self.submit_data
        )
        self.submit_button.grid(row=2, column=0, pady=(10, 20), sticky="n")

        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="red"
        )
        self.status_label.grid(row=3, column=0, pady=(0, 10))

        self.display_current_prompt()

    def display_current_prompt(self):
        if self.current_prompt_index < len(self.data_prompts):
            self.title_label.configure(text=self.data_prompts[self.current_prompt_index])
            self.entry_box.delete(0, tk.END)
            self.entry_box.focus_set()
            self.status_label.configure(text="")
        else:
            self.show_completion_screen()

    def submit_data(self):
        entered_text = self.entry_box.get().strip()

        if entered_text:
            if self.current_prompt_index == 0:
                # Set system_input_chosen here, after the first prompt
                self.system_input_chosen = entered_text.upper() 
                self.collected_data.append(entered_text) 

                if self.system_input_chosen == "YES 45TU FRONT SET(OG)":
                    self.data_prompts.extend(self.yes45tu_prompts)
                else:
                    self.data_prompts.extend(self.generic_prompts)
                
                self.current_prompt_index += 1
                self.display_current_prompt()
            else:
                self.collected_data.append(entered_text)
                self.current_prompt_index += 1
                self.display_current_prompt()
        else:
            self.status_label.configure(text="Please enter some text!")
            self.after(1500, lambda: self.status_label.configure(text=""))

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

    def call_excel_generator(self):
        expected_len = len(self.data_prompts)
        if len(self.collected_data) != expected_len:
            self.update_completion_status(f"Error: Incorrect number of inputs provided. Expected {expected_len}, got {len(self.collected_data)}.", "red")
            return

        try:
            system_input = self.collected_data[0]
            elevation_type = self.collected_data[1]
            total_count = int(self.collected_data[2])

            sqft_per_type = 0.0
            total_sqft = 0.0
            perimeter_ft = 0.0
            total_perimeter_ft = 0.0
            calculated_outputs = {}

            opening_width_inches = 0.0
            opening_height_inches = 0.0
            
            bays_wide = 0
            bays_tall = 0

            if not hasattr(self, 'system_input_chosen') or self.system_input_chosen == "":
                self.update_completion_status("Internal Error: System input choice not found.", "red")
                return

            if self.system_input_chosen == "YES 45TU FRONT SET(OG)":
                bays_wide = int(self.collected_data[3])
                bays_tall = int(self.collected_data[4])
                opening_width_inches = float(self.collected_data[5])
                opening_height_inches = float(self.collected_data[6])

                calculated_outputs = calculate_yes45tu_quantities(
                    bays_wide,
                    bays_tall,
                    total_count,
                    opening_width_inches,
                    opening_height_inches
                )
            else:
                opening_width_inches = float(self.collected_data[3])
                opening_height_inches = float(self.collected_data[4])

                self.update_completion_status(f"Warning: No specific calculations defined for '{system_input}'. Calculating basic geometry.", "orange")

            opening_width_feet = opening_width_inches / 12.0
            opening_height_feet = opening_height_inches / 12.0

            sqft_per_type = calculate_rectangle_area(opening_width_feet, opening_height_feet)
            total_sqft = sqft_per_type * total_count

            perimeter_ft = calculate_perimeter(opening_width_feet, opening_height_feet)
            total_perimeter_ft = perimeter_ft * total_count

        except (ValueError, IndexError) as e:
            self.update_completion_status(f"Error processing inputs: {e}. Please ensure correct data types were entered.", "red")
            return

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
            self.update_completion_status
        )


if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop()