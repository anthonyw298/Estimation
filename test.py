import json

# Use raw string to avoid unicode escape issues
with open(r'C:\Users\tonyw\OneDrive\Desktop\Estimation\projects\1_Elevations.json', 'r') as f:
    data = json.load(f)

# Extract and print Door descriptions
for key, value in data.items():
    if isinstance(value, dict):
        outputs = value.get("calculated_outputs", [])
        for item in outputs:
            if item.get("type") == "Door":
                print(item.get("description"))
