import re
import json
import os
from data.parts_data import parts_data
from data.part_number import PART_NUMBER_MAP

# Removed hardcoded EXTRA_MATERIALS_FILE

EPSILON = 1e-9

def parse_length_to_feet(length_str):
    """Converts various length formats to total feet."""
    if not isinstance(length_str, str) or not length_str.strip(): return 0.0
    length_str = length_str.replace('’', "'").replace('”', '"').replace('“', '"')
    feet, inches = 0.0, 0.0
    if (m := re.search(r"(\d+\.?\d*)\s*(ft|')", length_str, re.IGNORECASE)): feet = float(m.group(1))
    if (m := re.search(r"(\d+\.?\d*)\s*(in|\")", length_str, re.IGNORECASE)): inches = float(m.group(1))
    if feet or inches: return feet + (inches / 12)
    if (m := re.search(r"(\d+\.?\d*)", length_str)): return float(m.group(1))
    return 0.0

def load_extra_materials(extra_materials_file): # Added extra_materials_file parameter
    """Load extra materials leftovers from JSON file."""
    if os.path.exists(extra_materials_file):
        try: return json.load(open(extra_materials_file, 'r'))
        except json.JSONDecodeError:
            save_extra_materials({}, extra_materials_file) # Pass the file path
            return {}
    return {}

def save_extra_materials(materials, extra_materials_file): # Added extra_materials_file parameter
    """Save extra materials leftovers to JSON file."""
    try: json.dump(materials, open(extra_materials_file, 'w'), indent=4)
    except IOError as e: print(f"Error: Could not save {extra_materials_file}: {e}")

def get_price_by_part(part_number, requested_qty, current_extra_materials=None, summary=False, group=False, extra_materials_file="extra_materials.json"): # Added extra_materials_file parameter with default
    """Calculate price and material impact."""
    match = parts_data.get(part_number)
    if not match:
        return None, ("ft" if isinstance(requested_qty, (float, int)) and requested_qty % 1 != 0 else "pcs"), None

    list_price = float(match.get('List Price', 0.0))
    units_str = match.get('Units', "1 pcs.")
    length_str = match.get('Length', "")

    total_price, unit_type = 0.0, None
    
    part_extra_sim = {'quantity': 0, 'length_pieces': []}
    if not summary:
        if current_extra_materials is None: 
            current_extra_materials = load_extra_materials(extra_materials_file) # Pass the file path
        part_extra_sim = current_extra_materials.get(part_number, {'quantity': 0, 'length_pieces': []})

    material_impact_details = {
        'part_number': part_number, 'requested_qty': requested_qty, 'purchased_qty_or_length': 0.0,
        'leftover_generated_qty_or_length': 0.0, 'used_from_leftover_qty_or_length': 0.0,
        'cost_incurred': 0.0, 'type_processed_as': None
    }

    is_profile = (part_number in PART_NUMBER_MAP.get('profiles', {})) or group

    if is_profile:
        unit_type = "ft"
        min_purchase_length = parse_length_to_feet(length_str) or 1.0
        leftover_pieces_sim = sorted(list(part_extra_sim.get('length_pieces', [])))
        suitable_index = None
        if not summary:
            for i, piece_len in enumerate(leftover_pieces_sim):
                if piece_len >= requested_qty - EPSILON:
                    suitable_index = i
                    break

        if suitable_index is not None:
            total_price = 0.0 
            material_impact_details['used_from_leftover_qty_or_length'] = requested_qty
        else:
            num_bundles_needed = int(-(-requested_qty // min_purchase_length))
            actual_purchased_length = num_bundles_needed * min_purchase_length
            total_price = list_price * num_bundles_needed
            leftover_piece = max(0.0, actual_purchased_length - requested_qty)
            
            material_impact_details['purchased_qty_or_length'] = actual_purchased_length
            if leftover_piece > EPSILON: material_impact_details['leftover_generated_qty_or_length'] = leftover_piece
            material_impact_details['cost_incurred'] = total_price
        
        material_impact_details['type_processed_as'] = 'profile'

    else: # Assumed accessory or simple item
        unit_type = "pcs"
        unit_count_per_bundle = 1
        if 'pc' in units_str.lower():
            try: unit_count_per_bundle = int(units_str.lower().split('pc')[0].strip()) or 1
            except ValueError: pass
        
        leftover_qty_sim = part_extra_sim.get('quantity', 0)
        used_from_existing_leftover = 0
        if not summary: used_from_existing_leftover = min(requested_qty, leftover_qty_sim)
        
        remaining_needed_qty = requested_qty - used_from_existing_leftover
        actual_purchased_qty, excess_qty_from_new_purchase = 0, 0
        total_price = 0.0

        if remaining_needed_qty > 0:
            num_bundles_needed = (remaining_needed_qty + unit_count_per_bundle - 1) // unit_count_per_bundle
            actual_purchased_qty = num_bundles_needed * unit_count_per_bundle
            total_price = list_price * num_bundles_needed
            excess_qty_from_new_purchase = actual_purchased_qty - remaining_needed_qty
        
        material_impact_details['used_from_leftover_qty_or_length'] = used_from_existing_leftover
        material_impact_details['purchased_qty_or_length'] = actual_purchased_qty
        material_impact_details['leftover_generated_qty_or_length'] = excess_qty_from_new_purchase
        material_impact_details['cost_incurred'] = total_price
        material_impact_details['type_processed_as'] = 'accessory'

    return (total_price, unit_type, None) if summary else (total_price, unit_type, material_impact_details)


def apply_material_impact_to_extra_materials(material_impact_details, extra_materials_file="extra_materials.json"): # Added extra_materials_file parameter with default
    """Applies a single item's material impact to the extra_materials.json file."""
    if not material_impact_details or material_impact_details.get('part_number') == "N/A - Manual": return

    extra_materials = load_extra_materials(extra_materials_file) # Pass the file path
    part_number = material_impact_details.get('part_number')
    type_processed_as = material_impact_details.get('type_processed_as')
    if not part_number: return

    part_extra = extra_materials.get(part_number, {'quantity': 0, 'length_pieces': []})

    if type_processed_as == 'profile':
        used_from_leftover_qty_or_length = material_impact_details.get('used_from_leftover_qty_or_length', 0.0)
        if used_from_leftover_qty_or_length > EPSILON:
            temp_leftovers = sorted(list(part_extra.get('length_pieces', [])))
            consumed = False
            for i, piece_len in enumerate(temp_leftovers):
                if piece_len >= used_from_leftover_qty_or_length - EPSILON:
                    remaining_after_use = piece_len - used_from_leftover_qty_or_length
                    temp_leftovers.pop(i)
                    if remaining_after_use > EPSILON: temp_leftovers.append(remaining_after_use)
                    consumed = True
                    break
            if not consumed: print(f"⚠️ Warning: Could not find suitable leftover piece to consume {used_from_leftover_qty_or_length:.4f} for {part_number}.")
            part_extra['length_pieces'] = temp_leftovers
            part_extra['quantity'] = 0.0

        leftover_generated_qty_or_length = material_impact_details.get('leftover_generated_qty_or_length', 0.0)
        if leftover_generated_qty_or_length > EPSILON:
            part_extra.setdefault('length_pieces', []).append(leftover_generated_qty_or_length)
        part_extra['length_pieces'].sort()

    elif type_processed_as == 'accessory':
        current_qty = part_extra.get('quantity', 0)
        net_change = material_impact_details.get('leftover_generated_qty_or_length', 0.0) - material_impact_details.get('used_from_leftover_qty_or_length', 0.0)
        part_extra['quantity'] = round(current_qty + net_change, 4)
        part_extra['quantity'] = max(0, part_extra['quantity'])
        part_extra['length_pieces'] = []

    extra_materials[part_number] = part_extra
    save_extra_materials(extra_materials, extra_materials_file) # Pass the file path


def apply_material_impact_to_extra_materials_in_memory(materials_dict, material_impact_details):
    """Applies a single item's material impact to a provided materials dictionary in memory."""
    if not material_impact_details or material_impact_details.get('part_number') == "N/A - Manual": return

    part_number = material_impact_details.get('part_number')
    type_processed_as = material_impact_details.get('type_processed_as')
    if not part_number: return

    part_extra = materials_dict.get(part_number, {'quantity': 0, 'length_pieces': []})

    if type_processed_as == 'profile':
        used_from_leftover_qty_or_length = material_impact_details.get('used_from_leftover_qty_or_length', 0.0)
        if used_from_leftover_qty_or_length > EPSILON:
            temp_leftovers = sorted(list(part_extra.get('length_pieces', [])))
            consumed = False
            for i, piece_len in enumerate(temp_leftovers):
                if piece_len >= used_from_leftover_qty_or_length - EPSILON:
                    remaining_after_use = piece_len - used_from_leftover_qty_or_length
                    temp_leftovers.pop(i)
                    if remaining_after_use > EPSILON: temp_leftovers.append(remaining_after_use)
                    consumed = True
                    break
            if not consumed: print(f"⚠️ Warning (in-memory): Could not find suitable leftover piece to consume {used_from_leftover_qty_or_length:.4f} for {part_number}.")
            part_extra['length_pieces'] = temp_leftovers
            part_extra['quantity'] = 0.0

        leftover_generated_qty_or_length = material_impact_details.get('leftover_generated_qty_or_length', 0.0)
        if leftover_generated_qty_or_length > EPSILON:
            part_extra.setdefault('length_pieces', []).append(leftover_generated_qty_or_length)
        part_extra['length_pieces'].sort()

    elif type_processed_as == 'accessory':
        current_qty = part_extra.get('quantity', 0)
        net_change = material_impact_details.get('leftover_generated_qty_or_length', 0.0) - material_impact_details.get('used_from_leftover_qty_or_length', 0.0)
        part_extra['quantity'] = round(current_qty + net_change, 4)
        part_extra['quantity'] = max(0, part_extra['quantity'])
        part_extra['length_pieces'] = []

    materials_dict[part_number] = part_extra


def reverse_material_impact(elevation_material_impacts, extra_materials_file="extra_materials.json"): # Added extra_materials_file parameter with default
    """Reverses the material impact of a deleted elevation on extra_materials.json."""
    if not elevation_material_impacts: return

    extra_materials = load_extra_materials(extra_materials_file) # Pass the file path

    for impact in elevation_material_impacts:
        part_number = impact.get('part_number')
        type_processed_as = impact.get('type_processed_as')

        if not part_number or part_number == "N/A - Manual": continue

        part_extra = extra_materials.get(part_number, {'quantity': 0, 'length_pieces': []})

        leftover_generated_qty_or_length = impact.get('leftover_generated_qty_or_length', 0.0)
        used_from_leftover_qty_or_length = impact.get('used_from_leftover_qty_or_length', 0.0)
        
        if type_processed_as == 'profile':
            if leftover_generated_qty_or_length > EPSILON:
                removed = False
                for i, piece_len in enumerate(part_extra.get('length_pieces', [])):
                    if abs(piece_len - leftover_generated_qty_or_length) < EPSILON:
                        part_extra['length_pieces'].pop(i)
                        removed = True
                        break
                if not removed: print(f"⚠️ Warning: Generated leftover '{leftover_generated_qty_or_length:.4f} ft' for {part_number} not found in current inventory for reversal.")

            if used_from_leftover_qty_or_length > EPSILON:
                part_extra.setdefault('length_pieces', []).append(used_from_leftover_qty_or_length)
            
            part_extra['length_pieces'].sort()
            part_extra['quantity'] = 0.0

        elif type_processed_as == 'accessory':
            current_qty = part_extra.get('quantity', 0)
            reverse_net_change = used_from_leftover_qty_or_length - leftover_generated_qty_or_length
            part_extra['quantity'] = max(0, round(current_qty + reverse_net_change, 4))
            part_extra['length_pieces'] = []

        extra_materials[part_number] = part_extra
        
    save_extra_materials(extra_materials, extra_materials_file) # Pass the file path

def get_unit_price_by_part(part_number, extra_materials_file="extra_materials.json"): # Added extra_materials_file parameter with default
    """
    Retrieves the base list price per unit (foot for profiles, piece for accessories)
    for a given part number from parts_data.
    """
    match = parts_data.get(part_number)
    if not match:
        return None, None

    list_price = match.get('List Price', 0)
    units_str = match.get('Units', None)
    unit_count = 1

    if isinstance(units_str, str):
        units_lower = units_str.lower().strip()
        if 'pcs' in units_lower or 'pc' in units_lower:
            try:
                unit_count = int(units_lower.split('pc')[0].strip())
            except Exception:
                unit_count = 1

    if unit_count > 1:
        list_price /= unit_count

    length_str = match.get('Length', None)
    length_ft = parse_length_to_feet(length_str)

    unit_type = "pcs"

    if length_ft > 1:
        list_price /= length_ft
        unit_type = "ft"

    return list_price, unit_type