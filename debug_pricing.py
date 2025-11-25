from utils.pricing import get_unit_price_by_part, get_price_by_part
from data.parts_data import parts_data

part = "BE9-2513"
print(f"Testing part: {part}")
print(f"Raw data: {parts_data.get(part)}")

unit_price, unit_type = get_unit_price_by_part(part, finish="Clear")
print(f"Unit Price (Clear): {unit_price}, Type: {unit_type}")

total, unit, details = get_price_by_part(part, 16.0, finish="Clear", summary=False)
print(f"Total Price (16ft): {total}, Unit: {unit}")
print(f"Details: {details}")

