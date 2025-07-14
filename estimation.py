import openpyxl
from openpyxl.utils import get_column_letter

# Create a new Excel workbook and select the active worksheet
wb = openpyxl.Workbook()
ws = wb.active

# Get system input; if it doesn't match, create empty file and exit
system = input("System Input (YES 45TU Front Set(OG) or other): ")
if system != "YES 45TU Front Set(OG)":
    wb.save("output.xlsx")
    print("System not matched. Empty file created.")
    exit()

# Collect user inputs for the calculation
ElevationType = input("Enter Elevation Type: ")
TotalCount = int(input("Enter Total Count: "))
bays_wide = int(input("Enter # Bays Wide: "))
bays_tall = int(input("Enter # Bays Tall: "))
Opening_Width = float(input("Enter Opening Width: "))
Opening_Height = float(input("Enter Opening Height: "))
SqFt_per_type = float(input("Enter Sq Ft per Type: "))
Total_SqFt = float(input("Enter Total Sq Ft: "))
Perimeter_ft = float(input("Enter Perimeter Ft: "))
Total_Perimeter_ft = float(input("Enter Total Perimeter Ft: "))

# Define input and output headers
inputs = [
    "System Input", "Elevation Type", "Total Count",
    "# Bays Wide", "# Bays Tall",
    "Opening Width", "Opening Height",
    "Sq Ft per Type", "Total Sq Ft", "Perimeter Ft", "Total Perimeter Ft"
]

outputs = [
    "Total Gasket (Ft)", "End Dam", "Water Deflector", "Assembly Screw",
    "Sill Flash Screw", "End Dam Screw", "Setting Block Chair",
    "Side Block", "Setting Block", "Anti Walk Block Deep Pocket",
    "Anti Walk Block Shallow Pocket", "Setting Block (Int. Horizontal)",
    "Jamb Ft (V)", "Sill Ft (H)", "Flush Filler (V)",
    "Int. Vertical", "OG Int. Horizontal", "OG Head (H)", "Sill Flashing (H)"
]

# Combine input and output headers and write them to the first row
headers = inputs + outputs
for idx, header in enumerate(headers):
    ws[f"{get_column_letter(idx + 1)}1"] = header

# Prepare input values for row 2
input_values = [
    system, ElevationType, TotalCount,
    bays_wide, bays_tall, Opening_Width, Opening_Height,
    SqFt_per_type, Total_SqFt, Perimeter_ft, Total_Perimeter_ft
]

# Calculate output values directly and store them in a list
output_values = [
    (((bays_wide * 4 * Opening_Height) + (bays_tall * 4 * Opening_Width)) * TotalCount) / 12,  # Total Gasket (Ft)
    2 * TotalCount,  # End Dam
    2 * bays_wide * TotalCount,  # Water Deflector
    ((bays_wide * 8) + (((bays_tall - 1) * 6) * bays_wide)) * TotalCount,  # Assembly Screw
    3 * bays_wide * TotalCount,  # Sill Flash Screw
    4 * TotalCount,  # End Dam Screw
    2 * bays_wide,  # Setting Block Chair
    (bays_wide - 1) * bays_tall * TotalCount,  # Side Block
    2 * bays_wide * TotalCount,  # Setting Block
    2 * bays_tall * TotalCount,  # Anti Walk Block Deep Pocket
    (bays_wide - 1) * bays_tall * TotalCount,  # Anti Walk Block Shallow Pocket
    2 * bays_wide * TotalCount,  # Setting Block (Int. Horizontal)
    (2 * Opening_Height) / 12 * TotalCount,  # Jamb Ft (V)
    (Opening_Width / 12) * TotalCount,  # Sill Ft (H)
    ((bays_wide - 1) * TotalCount * Opening_Height) / 12,  # Flush Filler (V)
    ((bays_wide - 1) * TotalCount * Opening_Height) / 12,  # Int. Vertical (same as Flush Filler)
    (Opening_Width / 12) * TotalCount,  # OG Int. Horizontal
    (Opening_Width / 12) * TotalCount,  # OG Head (H)
    (Opening_Width / 12) * TotalCount   # Sill Flashing (H)
]

# Combine inputs and outputs and write them to row 2
all_values = input_values + output_values
for idx, val in enumerate(all_values):
    ws[f"{get_column_letter(idx + 1)}2"] = val

# Auto-size each column based on max length of header or data
for idx, header in enumerate(headers, 1):
    col_letter = get_column_letter(idx)
    ws.column_dimensions[col_letter].width = max(len(header), len(str(ws[f"{col_letter}2"].value))) + 2

# Save workbook
wb.save("output.xlsx")
print("Excel file saved as 'output.xlsx'")
