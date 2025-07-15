import customtkinter as ctk
import tkinter as tk
from estimation import generate_excel_report # Import the function

class App(ctk.CTk):
    """
    A CustomTkinter application for collecting data points for Excel calculations
    one at a time, then performing calculations and generating an Excel file.
    """
    def __init__(self):
        super().__init__()

        # --- Window Configuration ---
        self.title("Excel Data Entry Application")
        self.geometry("500x350") # Increased size for more prompts
        self.resizable(False, False) # Prevent resizing for a consistent look

        # Configure grid for centering the main frame
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Data Management ---
        # List of prompts for data collection, matching the Excel script's inputs
        self.data_prompts = [
            "Enter System Input (YES 45TU Front Set(OG) or other):",
            "Enter Elevation Type:",
            "Enter Total Count (integer):",
            "Enter # Bays Wide (integer):",
            "Enter # Bays Tall (integer):",
            "Enter Opening Width (float):",
            "Enter Opening Height (float):",
            "Enter Sq Ft per Type (float):",
            "Enter Total Sq Ft (float):",
            "Enter Perimeter Ft (float):",
            "Enter Total Perimeter Ft (float):"
        ]
        self.current_prompt_index = 0 # Tracks which prompt is currently active
        self.collected_data = []      # Stores all data entered by the user

        # --- UI Elements ---
        # Main frame to hold all widgets, centered and with padding
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        # Configure grid within the main frame for centering its contents
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1) # For the title label
        self.main_frame.grid_rowconfigure(1, weight=1) # For the entry box
        self.main_frame.grid_rowconfigure(2, weight=1) # For the submit button

        # Title Label: Displays the current prompt
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="", # Text will be set dynamically
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white",
            wraplength=400 # Allow text to wrap for longer prompts
        )
        self.title_label.grid(row=0, column=0, pady=(20, 10), sticky="s")

        # Entry Box: Where the user types their input
        self.entry_box = ctk.CTkEntry(
            self.main_frame,
            width=300,
            placeholder_text="Type here..."
        )
        self.entry_box.grid(row=1, column=0, pady=10, sticky="n")

        # Submit Button: Triggers data collection and advances to the next prompt
        self.submit_button = ctk.CTkButton(
            self.main_frame,
            text="Submit",
            command=self.submit_data
        )
        self.submit_button.grid(row=2, column=0, pady=(10, 20), sticky="n")

        # Status Label: For displaying temporary messages (e.g., "Please enter text!")
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="red"
        )
        self.status_label.grid(row=3, column=0, pady=(0, 10))


        # --- Initial Display ---
        self.display_current_prompt() # Show the first prompt when the app starts

    def display_current_prompt(self):
        """
        Updates the UI to show the current data prompt or
        transitions to the completion screen if all data is collected.
        """
        if self.current_prompt_index < len(self.data_prompts):
            self.title_label.configure(text=self.data_prompts[self.current_prompt_index])
            self.entry_box.delete(0, tk.END) # Clear any text from the previous entry
            self.entry_box.focus_set()       # Set keyboard focus to the entry box for convenience
            self.status_label.configure(text="") # Clear any previous status messages
        else:
            # All prompts have been answered, show the completion screen
            self.show_completion_screen()

    def submit_data(self):
        """
        Handles the submission of data from the entry box.
        Collects the data, clears the entry, and moves to the next prompt.
        """
        entered_text = self.entry_box.get().strip() # Get text and remove leading/trailing whitespace

        if entered_text:
            self.collected_data.append(entered_text) # Add the entered text to our list
            self.current_prompt_index += 1           # Move to the next prompt
            self.display_current_prompt()            # Update the UI
        else:
            # Provide feedback if the entry box is empty
            self.status_label.configure(text="Please enter some text!")
            # Revert the message after a short delay
            self.after(1500, lambda: self.status_label.configure(text=""))

    def show_completion_screen(self):
        """
        Displays a final screen indicating that all data has been collected.
        Also calls the `generate_excel_report` function from excel_calculator.py.
        """
        # Remove all existing widgets from the main frame
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        # Update the main frame's grid configuration for the completion screen
        self.main_frame.grid_rowconfigure(0, weight=1) # Make the completion label centered
        self.main_frame.grid_rowconfigure(1, weight=0)
        self.main_frame.grid_rowconfigure(2, weight=0)
        self.main_frame.grid_rowconfigure(3, weight=0)

        # Create a new title label for the completion screen
        self.completion_label = ctk.CTkLabel(
            self.main_frame,
            text="Processing data and generating Excel file...",
            font=ctk.CTkFont(size=18, weight="bold"),
            wraplength=300
        )
        self.completion_label.grid(row=0, column=0, pady=30, sticky="nsew")
        
        # Call the external function, passing collected data and a callback for UI updates
        # Use self.after to allow the UI to update with "Processing..." before blocking
        self.after(100, lambda: self.call_excel_generator())

    def update_completion_status(self, message: str, color: str):
        """
        Callback function to update the completion status label in the UI.
        """
        self.completion_label.configure(text=message, text_color=color)

    def call_excel_generator(self):
        """
        Prepares data and calls the generate_excel_report function from excel_calculator.py.
        """
        # Ensure correct number of arguments
        if len(self.collected_data) != len(self.data_prompts):
            self.update_completion_status("Error: Incorrect number of inputs provided.", "red")
            return

        # Map collected data to variables and perform type conversions
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
            self.update_completion_status(f"Error: Invalid number format. Please check your inputs. ({e})", "red")
            return

        # Call the external function from excel_calculator.py
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
            self.update_completion_status # Pass the UI update callback
        )


# --- Application Entry Point ---
if __name__ == "__main__":
    # Set the appearance mode (Light, Dark, System)
    ctk.set_appearance_mode("Dark")
    # Set the default color theme (blue, dark-blue, green)
    ctk.set_default_color_theme("blue")

    app = App()
    app.mainloop() # Start the CustomTkinter event loop
