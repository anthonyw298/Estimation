import re
import json
import os
import math # Import math for ceil
from data.parts_data import parts_data
from data.part_number import PART_NUMBER_MAP

EPSILON = 1e-9

def parse_length_to_feet(length_str):
    """Converts various length formats to total feet."""
    if not isinstance(length_str, str) or not length_str.strip(): return 0.0
    length_str = length_str.replace('’', "'").replace('”', '"').replace('“', '"')
    feet, inches = 0.0, 0.0
    # Search for feet (e.g., "10ft", "10'")
    if (m := re.search(r"(\d+\.?\d*)\s*(ft|')", length_str, re.IGNORECASE)): feet = float(m.group(1))
    # Search for inches (e.g., "6in", "6\"")
    if (m := re.search(r"(\d+\.?\d*)\s*(in|\")", length_str, re.IGNORECASE)): inches = float(m.group(1))
    
    # If feet or inches were found, calculate total
    if feet or inches: return feet + (inches / 12)
    
    # Fallback: if no units, assume the number itself is in feet
    if (m := re.search(r"(\d+\.?\d*)", length_str)): return float(m.group(1))
    
    return 0.0

def load_extra_materials(extra_materials_file):
    """Load extra materials leftovers from JSON file."""
    if os.path.exists(extra_materials_file):
        try: return json.load(open(extra_materials_file, 'r'))
        except json.JSONDecodeError:
            save_extra_materials({}, extra_materials_file)
            return {}
    return {}

def save_extra_materials(materials, extra_materials_file):
    """Save extra materials leftovers to JSON file."""
    try: json.dump(materials, open(extra_materials_file, 'w'), indent=4)
    except IOError as e: print(f"Error: Could not save {extra_materials_file}: {e}")

def get_unit_price_by_part(part_number, finish=None, extra_materials_file="extra_materials.json"):
    """
    Retrieves the base list price per unit (foot for profiles/items with length, piece for accessories)
    for a given part number from parts_data, considering finish for profiles.
    """
    match = parts_data.get(part_number)
    if not match:
        return None, None

    list_price_raw = match.get('List Price', 0) # Can be float or a list [clear, black, paint]
    units_str = match.get('Units', None)
    length_str = match.get('Length', None)

    unit_count = 1
    if isinstance(units_str, str):
        units_lower = units_str.lower().strip()
        if 'pcs' in units_lower or 'pc' in units_lower:
            try:
                # Extract number before 'pc' for units like "5 pcs."
                num_part = units_lower.split('pc')[0].strip()
                if num_part: # Ensure it's not empty string (e.g., for "1pc." or "pcs")
                    unit_count = int(num_part)
            except ValueError: # Handle cases like just "pcs" or invalid formats
                unit_count = 1 

    list_price_effective = 0.0 # This will be the price for a 'single bundle/piece' before final unit division (by length or unit_count_per_bundle)
    unit_type = "pcs" # Default unit type

    is_profile = (part_number in PART_NUMBER_MAP.get('profiles', []))

    if is_profile:
        unit_type = "ft"
        # 1. Select base price based on finish (if list)
        if isinstance(list_price_raw, list) and len(list_price_raw) == 3:
            if finish and finish.lower() == 'clear':
                list_price_effective = float(list_price_raw[0])
            elif finish and finish.lower() == 'black':
                list_price_effective = float(list_price_raw[1])
            elif finish and finish.lower() == 'paint':
                list_price_effective = float(list_price_raw[2])
            else:
                # Default to clear price if finish is not specified or recognized
                list_price_effective = float(list_price_raw[0])
        else:
            # Fallback if List Price is not a list for a profile (e.g., old data or single price profile)
            try:
                list_price_effective = float(list_price_raw)
            except (TypeError, ValueError):
                list_price_effective = 0.0

        # 2. Apply the finish-based multiplier for profiles
        finish_multiplier = 1.0
        if finish and finish.lower() == 'black':
            finish_multiplier = 1.1
        elif finish and finish.lower() == 'paint':
            finish_multiplier = 1.2
        list_price_effective *= finish_multiplier

        # 3. Divide by length to get per-foot price for profiles
        length_ft = parse_length_to_feet(length_str)
        if length_ft > EPSILON: # Use EPSILON for floating point comparison against zero
            list_price_effective /= length_ft
        else:
            print(f"Warning: Profile '{part_number}' has zero or invalid length. Unit price might be incorrect.")

    else: # Not a profile (accessory or other non-profile item)
        # For non-profiles, the List Price is usually a single value
        try:
            list_price_effective = float(list_price_raw)
        except (TypeError, ValueError):
            list_price_effective = 0.0

        # Apply unit_count division (e.g., if price is for a pack of 5, divide by 5 to get per-piece)
        if unit_count > 1:
            list_price_effective /= unit_count

        # Apply the length-based division if length is significant (your heuristic: > 1 foot)
        # This handles items like E2-0052 which have a total price for a long length
        length_ft = parse_length_to_feet(length_str)
        if length_ft > 1: # This is your crucial heuristic for when to price per foot for non-profiles
            list_price_effective /= length_ft
            unit_type = "ft" # Change unit type to feet as we now have a per-foot price

    return list_price_effective, unit_type


def get_price_by_part(part_number, requested_qty, finish=None, current_extra_materials=None, summary=False, group=False, extra_materials_file="extra_materials.json"):
    """
    Calculate price and material impact, considering finish for profiles.
    
    Args:
        part_number (str): The part number.
        requested_qty (float/int): The quantity requested.
        finish (str, optional): The finish type ('clear', 'black', 'paint'). Only relevant for profiles.
        current_extra_materials (dict, optional): In-memory extra materials state.
        summary (bool): If True, only return price and unit type, no material impact details.
        group (bool): If True, forces profile-like behavior for pricing.
        extra_materials_file (str): Path to the extra materials JSON file.
    
    Returns:
        tuple: (total_price, unit_type, material_impact_details) or (total_price, unit_type, None) if summary.
    """
    match = parts_data.get(part_number)
    if not match:
        return None, ("ft" if isinstance(requested_qty, (float, int)) and requested_qty % 1 != 0 else "pcs"), None

    units_str = match.get('Units', "1 pcs.")
    length_str = match.get('Length', "")

    # *** MAJOR FIX: Use get_unit_price_by_part to get the correct unit price ***
    unit_price, unit_type = get_unit_price_by_part(part_number, finish, extra_materials_file)
    
    if unit_price is None:
        # If get_unit_price_by_part failed, propagate None
        return None, unit_type, None # unit_type will reflect default or what get_unit_price_by_part tried to determine

    total_price = 0.0
    
    # Construct the key for extra materials based on part number and finish (for profiles)
    extra_materials_key = part_number
    # Determine if this part is treated as a profile for inventory purposes
    is_profile_for_inventory = (part_number in PART_NUMBER_MAP.get('profiles', [])) or group
    if is_profile_for_inventory and finish:
        extra_materials_key = f"{part_number}-{finish.lower()}"

    part_extra_sim = {'quantity': 0, 'length_pieces': []}
    if not summary:
        if current_extra_materials is None: 
            current_extra_materials = load_extra_materials(extra_materials_file)
        part_extra_sim = current_extra_materials.get(extra_materials_key, {'quantity': 0, 'length_pieces': []})

    material_impact_details = {
        'part_number': part_number,
        'requested_qty': requested_qty,
        'purchased_qty_or_length': 0.0,
        'leftover_generated_qty_or_length': 0.0,
        'used_from_leftover_qty_or_length': 0.0,
        'cost_incurred': 0.0,
        'type_processed_as': None,
        'finish': finish # Store the finish in material impact details
    }

    if is_profile_for_inventory: # This logic handles profiles and length-based items (like E2-0052 now that unit_type is 'ft')
        unit_type = "ft" # Confirm unit type for reporting
        min_purchase_length = parse_length_to_feet(length_str) or 1.0
        leftover_pieces_sim = sorted(list(part_extra_sim.get('length_pieces', [])))
        suitable_index = None
        
        if not summary:
            for i, piece_len in enumerate(leftover_pieces_sim):
                if piece_len >= requested_qty - EPSILON:
                    suitable_index = i
                    break

        if suitable_index is not None:
            # Material is taken from existing leftovers
            total_price = 0.0 # No new cost incurred
            material_impact_details['used_from_leftover_qty_or_length'] = requested_qty
            # The piece used from leftover will be adjusted/removed in apply_material_impact
        else:
            # No suitable leftover found, purchase new material
            num_bundles_needed = math.ceil(requested_qty / min_purchase_length)
            
            actual_purchased_length = num_bundles_needed * min_purchase_length
            # *** Use the 'unit_price' obtained from get_unit_price_by_part ***
            total_price = unit_price * actual_purchased_length # Calculate total price using the per-foot unit_price
            
            leftover_piece = max(0.0, actual_purchased_length - requested_qty)
            
            material_impact_details['purchased_qty_or_length'] = actual_purchased_length
            if leftover_piece > EPSILON: 
                material_impact_details['leftover_generated_qty_or_length'] = leftover_piece
            material_impact_details['cost_incurred'] = total_price
        
        material_impact_details['type_processed_as'] = 'profile' # Even if it's E2-0052, it's processed like a profile for inventory

    else: # Assumed accessory or simple item (piece-based)
        # Unit type for these items is already set correctly by get_unit_price_by_part, but confirm for consistency
        # If get_unit_price_by_part returned 'ft' for a non-profile (e.g., E2-0052 if it wasn't caught by `is_profile_for_inventory` based on `group`),
        # this block needs to be careful. However, with `is_profile_for_inventory` covering E2-0052, this `else` is truly for pieces.
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
            num_bundles_needed = math.ceil(remaining_needed_qty / unit_count_per_bundle)
            actual_purchased_qty = num_bundles_needed * unit_count_per_bundle
            # *** Use the 'unit_price' obtained from get_unit_price_by_part ***
            total_price = unit_price * actual_purchased_qty # Calculate total price using the per-piece unit_price
            
            excess_qty_from_new_purchase = actual_purchased_qty - remaining_needed_qty
        
        material_impact_details['used_from_leftover_qty_or_length'] = used_from_existing_leftover
        material_impact_details['purchased_qty_or_length'] = actual_purchased_qty
        material_impact_details['leftover_generated_qty_or_length'] = excess_qty_from_new_purchase
        material_impact_details['cost_incurred'] = total_price
        material_impact_details['type_processed_as'] = 'accessory'

    return (total_price, unit_type, None) if summary else (total_price, unit_type, material_impact_details)


def apply_material_impact_to_extra_materials(material_impact_details, extra_materials_file="extra_materials.json"):
    """Applies a single item's material impact to the extra_materials.json file."""
    if not material_impact_details or material_impact_details.get('part_number') == "N/A - Manual": return

    extra_materials = load_extra_materials(extra_materials_file)
    part_number = material_impact_details.get('part_number')
    type_processed_as = material_impact_details.get('type_processed_as')
    finish = material_impact_details.get('finish') # Get finish from impact details
    if not part_number: return

    # Construct the key for extra materials based on part number and finish (for profiles)
    extra_materials_key = part_number
    if type_processed_as == 'profile' and finish: # Only profiles get finish in their inventory key
        extra_materials_key = f"{part_number}-{finish.lower()}"

    part_extra = extra_materials.get(extra_materials_key, {'quantity': 0, 'length_pieces': []})

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
            if not consumed: print(f"⚠️ Warning: Could not find suitable leftover piece to consume {used_from_leftover_qty_or_length:.4f} for {part_number} ({finish}).")
            part_extra['length_pieces'] = temp_leftovers
            part_extra['quantity'] = 0.0 # Profiles only use length_pieces, quantity is effectively 0 for whole pieces

        leftover_generated_qty_or_length = material_impact_details.get('leftover_generated_qty_or_length', 0.0)
        if leftover_generated_qty_or_length > EPSILON:
            part_extra.setdefault('length_pieces', []).append(leftover_generated_qty_or_length)
        part_extra['length_pieces'].sort()

    elif type_processed_as == 'accessory':
        current_qty = part_extra.get('quantity', 0)
        net_change = material_impact_details.get('leftover_generated_qty_or_length', 0.0) - material_impact_details.get('used_from_leftover_qty_or_length', 0.0)
        part_extra['quantity'] = round(current_qty + net_change, 4)
        part_extra['quantity'] = max(0, part_extra['quantity'])
        part_extra['length_pieces'] = [] # Accessories only use quantity, length_pieces is effectively empty

    extra_materials[extra_materials_key] = part_extra
    save_extra_materials(extra_materials, extra_materials_file)


def apply_material_impact_to_extra_materials_in_memory(materials_dict, material_impact_details):
    """Applies a single item's material impact to a provided materials dictionary in memory."""
    if not material_impact_details or material_impact_details.get('part_number') == "N/A - Manual": return

    part_number = material_impact_details.get('part_number')
    type_processed_as = material_impact_details.get('type_processed_as')
    finish = material_impact_details.get('finish') # Get finish from impact details
    if not part_number: return

    # Construct the key for extra materials based on part number and finish (for profiles)
    extra_materials_key = part_number
    if type_processed_as == 'profile' and finish: # Only profiles get finish in their inventory key
        extra_materials_key = f"{part_number}-{finish.lower()}"

    part_extra = materials_dict.get(extra_materials_key, {'quantity': 0, 'length_pieces': []})

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
            if not consumed: print(f"⚠️ Warning (in-memory): Could not find suitable leftover piece to consume {used_from_leftover_qty_or_length:.4f} for {part_number} ({finish}).")
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

    materials_dict[extra_materials_key] = part_extra


def reverse_material_impact(elevation_material_impacts, extra_materials_file="extra_materials.json"):
    """Reverses the material impact of a deleted elevation on extra_materials.json."""
    if not elevation_material_impacts: return

    extra_materials = load_extra_materials(extra_materials_file)

    for impact in elevation_material_impacts:
        part_number = impact.get('part_number')
        type_processed_as = impact.get('type_processed_as')
        finish = impact.get('finish') # Get finish from impact details

        if not part_number or part_number == "N/A - Manual": continue

        # Construct the key for extra materials based on part number and finish (for profiles)
        extra_materials_key = part_number
        if type_processed_as == 'profile' and finish:
            extra_materials_key = f"{part_number}-{finish.lower()}"

        part_extra = extra_materials.get(extra_materials_key, {'quantity': 0, 'length_pieces': []})

        leftover_generated_qty_or_length = impact.get('leftover_generated_qty_or_length', 0.0)
        used_from_leftover_qty_or_length = impact.get('used_from_leftover_qty_or_length', 0.0)
        
        if type_processed_as == 'profile':
            # When reversing, if a leftover was GENERATED, we remove it from inventory.
            if leftover_generated_qty_or_length > EPSILON:
                removed = False
                temp_leftovers = list(part_extra.get('length_pieces', [])) # Create a mutable copy
                for i, piece_len in enumerate(temp_leftovers):
                    if abs(piece_len - leftover_generated_qty_or_length) < EPSILON:
                        temp_leftovers.pop(i) # Remove the specific piece
                        removed = True
                        break
                if not removed: print(f"⚠️ Warning: Generated leftover '{leftover_generated_qty_or_length:.4f} ft' for {part_number} ({finish}) not found in current inventory for reversal.")
                part_extra['length_pieces'] = temp_leftovers # Update the original list

            # When reversing, if material was USED FROM leftover, we put it back into inventory.
            if used_from_leftover_qty_or_length > EPSILON:
                part_extra.setdefault('length_pieces', []).append(used_from_leftover_qty_or_length)
            
            part_extra['length_pieces'].sort()
            part_extra['quantity'] = 0.0 # Profiles only use length_pieces

        elif type_processed_as == 'accessory':
            current_qty = part_extra.get('quantity', 0)
            reverse_net_change = used_from_leftover_qty_or_length - leftover_generated_qty_or_length
            part_extra['quantity'] = max(0, round(current_qty + reverse_net_change, 4))
            part_extra['length_pieces'] = [] # Accessories only use quantity

        extra_materials[extra_materials_key] = part_extra
        
    save_extra_materials(extra_materials, extra_materials_file)