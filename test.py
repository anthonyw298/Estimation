from data.part_number import (PART_NUMBER_MAP)

value = PART_NUMBER_MAP.get("accessories", {}).get("Total Gasket (Ft)", "Not Found")
print(value)  # Outputs: value3
