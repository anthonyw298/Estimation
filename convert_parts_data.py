import json
import sys
import os

# Add the current directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data.parts_data import parts_data

# Convert to JSON
output_path = os.path.join('src', 'data', 'parts_data.json')
os.makedirs(os.path.dirname(output_path), exist_ok=True)

with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(parts_data, f, indent=2, ensure_ascii=False)

print(f"✅ Converted parts_data.py to {output_path}")

