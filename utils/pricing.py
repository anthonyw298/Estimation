import re
import json
import os
from data.parts_data import parts_data
from data.part_number import PART_NUMBER_MAP

EXTRA_MATERIALS_FILE = "extra_materials.json"

def parse_length_to_feet(length_str):
    """
    Converts various length formats to total feet.
    Examples: 8', 96", 8 ft, 8ft 6in.
    Returns 0.0 if input is invalid or empty.
    """
    if not isinstance(length_str, str) or not length_str.strip():
        return 0.0

    length_str = length_str.replace('’', "'").replace('”', '"').replace('“', '"')

    feet = 0.0
    inches = 0.0

    feet_match = re.search(r"(\d+\.?\d*)\s*(ft|')", length_str, re.IGNORECASE)
    if feet_match:
        feet = float(feet_match.group(1))

    inches_match = re.search(r"(\d+\.?\d*)\s*(in|\")", length_str, re.IGNORECASE)
    if inches_match:
        inches = float(inches_match.group(1))

    if feet or inches:
        return feet + (inches / 12)

    num_match = re.search(r"(\d+\.?\d*)", length_str)
    if num_match:
        return float(num_match.group(1))

    return 0.0

def load_extra_materials():
    """Load extra materials leftovers from JSON file."""
    if os.path.exists(EXTRA_MATERIALS_FILE):
        try:
            with open(EXTRA_MATERIALS_FILE, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            print(f"Warning: Could not decode {EXTRA_MATERIALS_FILE}: {e}. Attempting to reset to empty.")
            # If corrupted, reset the file to an empty dictionary
            save_extra_materials({})
            return {}
    return {}

def save_extra_materials(materials):
    """Save extra materials leftovers to JSON file."""
    try:
        with open(EXTRA_MATERIALS_FILE, 'w') as f:
            json.dump(materials, f, indent=4)
    except IOError as e:
        print(f"Error: Could not save {EXTRA_MATERIALS_FILE}: {e}")


def get_price_by_part(part_number, requested_qty, current_extra_materials=None, summary=False, group=False):
    """
    Calculate price and material impact.
    If `summary` is True, calculations do not use or modify extra_materials.json.
    `current_extra_materials` allows for simulation based on a given state when summary is False.
    `group` can be set to True to force length-based calculations for items not explicitly in profiles.
    Returns total_price, unit_type, and material_impact_details.
    """
    match = parts_data.get(part_number)
    if not match:
        # For unknown parts, return None for price and impact, but still a unit type if possible
        unit_type_from_str = "pcs"
        if isinstance(requested_qty, (float, int)):
            if requested_qty % 1 != 0: # Likely a length if not whole number
                unit_type_from_str = "ft"
        return None, unit_type_from_str, None # Return None for impact details if match not found

    list_price = float(match.get('List Price', 0.0))
    units_str = match.get('Units', "1 pcs.")
    length_str = match.get('Length', "")

    total_price = 0.0
    unit_type = None
    
    # Use provided extra materials for simulation, or load them if not provided
    # If summary is True, always start with effectively empty extra materials for pricing.
    if summary:
        part_extra_sim = {'quantity': 0, 'length_pieces': []}
    else:
        if current_extra_materials is None:
            current_extra_materials = load_extra_materials()
        part_extra_sim = current_extra_materials.get(part_number, {'quantity': 0, 'length_pieces': []})

    material_impact_details = {
        'part_number': part_number,
        'requested_qty': requested_qty,
        'purchased_qty_or_length': 0.0,
        'leftover_generated_qty_or_length': 0.0,
        'used_from_leftover_qty_or_length': 0.0,
        'cost_incurred': 0.0 # This cost is before applying any finish multiplier
    }

    # Determine if it's a profile (length-based) or accessory (piece-based) based on PART_NUMBER_MAP
    # The 'group' parameter now explicitly forces profile-like behavior if set to True.
    is_profile = (part_number in PART_NUMBER_MAP.get('profiles', {})) or group
    is_accessory = part_number in PART_NUMBER_MAP.get('accessories', {})

    if is_profile:
        unit_type = "ft"
        min_purchase_length = parse_length_to_feet(length_str)
        if min_purchase_length <= 0:
            min_purchase_length = 1.0

        leftover_pieces_sim = list(part_extra_sim.get('length_pieces', [])) # Use a copy for simulation
        leftover_pieces_sim.sort()

        suitable_index = None
        if not summary: # Only consider existing leftovers if not in summary mode
            for i, piece_len in enumerate(leftover_pieces_sim):
                if piece_len >= requested_qty:
                    suitable_index = i
                    break

        if suitable_index is not None:
            # Used from leftover: no new purchase cost incurred for this item
            total_price = 0.0 
            material_impact_details['used_from_leftover_qty_or_length'] = requested_qty
            
        else:
            # New purchase needed
            num_bundles_needed = int(-(-requested_qty // min_purchase_length)) # Ceiling division
            actual_purchased_length = num_bundles_needed * min_purchase_length
            total_price = list_price * num_bundles_needed

            leftover_piece = actual_purchased_length - requested_qty
            
            material_impact_details['purchased_qty_or_length'] = actual_purchased_length
            if leftover_piece > 0:
                material_impact_details['leftover_generated_qty_or_length'] = leftover_piece
            material_impact_details['cost_incurred'] = total_price

    elif is_accessory:
        unit_type = "pcs"
        unit_count_per_bundle = 1
        if 'pc' in units_str.lower():
            try:
                pcs_part = units_str.lower().split('pc')[0].strip()
                unit_count_per_bundle = int(pcs_part) if pcs_part else 1
            except ValueError:
                pass
        
        leftover_qty_sim = part_extra_sim.get('quantity', 0)
        
        used_from_existing_leftover = 0
        if not summary: # Only consider existing leftovers if not in summary mode
            used_from_existing_leftover = min(requested_qty, leftover_qty_sim)
        
        remaining_needed_qty = requested_qty - used_from_existing_leftover

        actual_purchased_qty = 0
        excess_qty_from_new_purchase = 0
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

    else:
        # If part is found in parts_data but not in profiles or accessories, treat as a simple item
        # with no material impact tracking.
        total_price = list_price * requested_qty
        unit_type_match = re.search(r'\d+\s*([a-zA-Z.]+)', units_str)
        unit_type = unit_type_match.group(1).strip() if unit_type_match else "pcs"
        return total_price, unit_type, None # No material impact details for these items

    # If in summary mode, do not return material impact details as they are not used or needed.
    if summary:
        return total_price, unit_type, None
    else:
        return total_price, unit_type, material_impact_details


def apply_material_impact_to_extra_materials(material_impact_details):
    """
    Applies a single item's material impact to the extra_materials.json file.
    This function *modifies* extra_materials.json.
    """
    if not material_impact_details or material_impact_details.get('part_number') == "N/A - Manual": # Skip manual items without PN
        return

    extra_materials = load_extra_materials()
    part_number = material_impact_details.get('part_number')
    if not part_number:
        return

    part_extra = extra_materials.get(part_number, {'quantity': 0, 'length_pieces': []})

    # Use a small epsilon for floating point comparisons
    EPSILON = 1e-9

    if part_number in PART_NUMBER_MAP['profiles']:
        # If material was used from an existing leftover, we need to find and consume it.
        used_from_leftover_qty_or_length = material_impact_details.get('used_from_leftover_qty_or_length', 0.0)
        if used_from_leftover_qty_or_length > EPSILON:
            temp_leftovers = list(part_extra['length_pieces'])
            temp_leftovers.sort() # Ensure sorted for consumption logic
            
            consumed = False
            # Find the smallest piece that can satisfy the request
            for i, piece_len in enumerate(temp_leftovers):
                if piece_len >= used_from_leftover_qty_or_length - EPSILON:
                    remaining_after_use = piece_len - used_from_leftover_qty_or_length
                    temp_leftovers.pop(i) # Remove the consumed piece
                    if remaining_after_use > EPSILON: # Add back any new leftover from this piece if significant
                        temp_leftovers.append(remaining_after_use)
                    consumed = True
                    break
            
            if not consumed:
                print(f"⚠️ Warning: Could not find suitable leftover piece to consume {used_from_leftover_qty_or_length:.4f} for {part_number} during apply_impact. This might indicate a discrepancy in leftover tracking or floating point issues.")
            
            part_extra['length_pieces'] = temp_leftovers

        # If a leftover was generated by this purchase, add it.
        leftover_generated_qty_or_length = material_impact_details.get('leftover_generated_qty_or_length', 0.0)
        if leftover_generated_qty_or_length > EPSILON:
            part_extra['length_pieces'].append(leftover_generated_qty_or_length)
        
        part_extra['length_pieces'].sort() # Keep sorted

    elif part_number in PART_NUMBER_MAP['accessories']:
        # Update current quantity based on consumption and generation
        current_qty = part_extra.get('quantity', 0)
        leftover_generated_qty_or_length = material_impact_details.get('leftover_generated_qty_or_length', 0.0)
        used_from_leftover_qty_or_length = material_impact_details.get('used_from_leftover_qty_or_length', 0.0)

        net_change = leftover_generated_qty_or_length - used_from_leftover_qty_or_length
        part_extra['quantity'] = round(current_qty + net_change, 4) # Round to avoid float issues
        part_extra['quantity'] = max(0, part_extra['quantity']) # Ensure no negative quantity

    extra_materials[part_number] = part_extra
    save_extra_materials(extra_materials)


def reverse_material_impact(elevation_material_impacts):
    """
    Reverses the material impact of a deleted elevation on extra_materials.json.
    Takes a list of impact dictionaries for a single elevation.
    """
    if not elevation_material_impacts:
        print("ℹ️ No material impacts to reverse for this elevation.")
        return

    extra_materials = load_extra_materials()
    print(f"🔄 Reversing material impact for deleted elevation. Initial extra_materials: {extra_materials}")

    EPSILON = 1e-9 # Define a small epsilon for floating point comparisons

    for impact in elevation_material_impacts:
        part_number = impact.get('part_number')
        if not part_number or part_number == "N/A - Manual": # Skip manual items without PN
            continue

        part_extra = extra_materials.get(part_number, {'quantity': 0, 'length_pieces': []})

        purchased_qty_or_length = impact.get('purchased_qty_or_length', 0.0)
        leftover_generated_qty_or_length = impact.get('leftover_generated_qty_or_length', 0.0)
        used_from_leftover_qty_or_length = impact.get('used_from_leftover_qty_or_length', 0.0)
        
        print(f"   Processing impact for {part_number}:")
        print(f"     - Purchased (orig): {purchased_qty_or_length:.4f}")
        print(f"     - Leftover Generated (orig): {leftover_generated_qty_or_length:.4f}")
        print(f"     - Used from Leftover (orig): {used_from_leftover_qty_or_length:.4f}")


        # Determine if it's a profile (length-based) or accessory (piece-based)
        if part_number in PART_NUMBER_MAP['profiles']:
            
            # If a leftover was generated by this specific purchase, remove it.
            if leftover_generated_qty_or_length > EPSILON:
                removed = False
                for i, piece_len in enumerate(part_extra['length_pieces']):
                    if abs(piece_len - leftover_generated_qty_or_length) < EPSILON: # Check for near equality
                        part_extra['length_pieces'].pop(i)
                        removed = True
                        print(f"   - Removed generated leftover '{leftover_generated_qty_or_length:.4f} ft' for {part_number}")
                        break
                if not removed:
                    print(f"   - Warning: Generated leftover '{leftover_generated_qty_or_length:.4f} ft' for {part_number} not found in current inventory for reversal. Cannot remove. This might indicate previous manual modification or a bug.")

            # If material was used from an existing leftover, "return" it to inventory.
            if used_from_leftover_qty_or_length > EPSILON:
                part_extra['length_pieces'].append(used_from_leftover_qty_or_length)
                print(f"   - Returned used material '{used_from_leftover_qty_or_length:.4f} ft' to leftovers for {part_number}")
            
            part_extra['length_pieces'].sort() # Keep sorted

        elif part_number in PART_NUMBER_MAP['accessories']:
            current_qty = part_extra.get('quantity', 0)
            
            # The reversal of the net change: used_from_leftover - leftover_generated
            reverse_net_change = used_from_leftover_qty_or_length - leftover_generated_qty_or_length
            part_extra['quantity'] = round(current_qty + reverse_net_change, 4) # Apply reversal and round
            part_extra['quantity'] = max(0, part_extra['quantity']) # Ensure no negative quantity

            print(f"   - Accessory {part_number}: Previous Qty: {current_qty:.4f}, Net Change to Reverse: {reverse_net_change:.4f}, New Qty: {part_extra['quantity']:.4f}")
            if abs(reverse_net_change) > EPSILON: # Only print if there was a meaningful change
                if reverse_net_change > 0:
                    print(f"   - Returned {used_from_leftover_qty_or_length:.4f} pcs used and/or removed {leftover_generated_qty_or_length:.4f} pcs generated for {part_number}.")
                else: # reverse_net_change < 0
                    print(f"   - Removed {leftover_generated_qty_or_length:.4f} pcs generated and/or returned {used_from_leftover_qty_or_length:.4f} pcs used for {part_number}.")
            else:
                print(f"   - No significant net change to accessory {part_number} quantity.")

        extra_materials[part_number] = part_extra
        
    save_extra_materials(extra_materials)
    print(f"✅ Material impact reversed. Final extra_materials: {extra_materials}")
