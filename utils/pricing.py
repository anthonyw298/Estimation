import re
import json
import os
import math # Import math for ceil
from data.parts_data import parts_data
from data.part_number import PART_NUMBER_MAP

EPSILON = 1e-9

# Parts that use bay width/height lists for waste optimization
# Horizontal parts: use custom bay widths
HORIZONTAL_BAY_PARTS = {"BE9-2514", "BE9-2515", "E9-2519"}
# Vertical parts: use height/2 split
VERTICAL_BAY_PARTS = {"E9-2512", "BE9-2511"}
BAY_WIDTH_PARTS = HORIZONTAL_BAY_PARTS | VERTICAL_BAY_PARTS

def _is_bay_width_part(part_number, requested_qty=None, description=None):
    """
    Check if a part uses bay width/height lists for waste optimization.
    For BE9-2513, both sill (horizontal) and jamb (vertical) use lists.
    Vertical parts (2513 jamb, 2512, 2511) use height/2 split.
    We can detect this by checking if requested_qty is a list or by description.
    """
    if part_number == "BE9-2513":
        # Both sill and jamb use lists now (sill uses bay widths, jamb uses height/2)
        if requested_qty is not None:
            return isinstance(requested_qty, list)
        # Fallback to description check if requested_qty not available
        return description and ("sill" in description.lower() or "jamb" in description.lower() or "vertical" in description.lower())
    
    # Check if it's a vertical or horizontal bay part
    if part_number in VERTICAL_BAY_PARTS:
        # Vertical parts always use lists (height/2 split)
        return requested_qty is None or isinstance(requested_qty, list)
    
    return part_number in HORIZONTAL_BAY_PARTS

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
            # Normalize finish input to handle cases where it might be None or formatted differently
            finish_norm = finish.lower() if finish else 'clear'
            
            if finish_norm == 'clear':
                list_price_effective = float(list_price_raw[0])
            elif finish_norm == 'black':
                list_price_effective = float(list_price_raw[1])
            elif finish_norm == 'paint':
                list_price_effective = float(list_price_raw[2])
            else:
                # Default to clear price if finish is not recognized
                list_price_effective = float(list_price_raw[0])
        else:
            # Fallback if List Price is not a list for a profile
            try:
                list_price_effective = float(list_price_raw)
            except (TypeError, ValueError):
                list_price_effective = 0.0

        # 2. Apply the finish-based multiplier for profiles
        # Only apply this if we haven't already selected a specific price from the list above
        # Logic check: if list_price_raw was a list, we already picked the correct price.
        # If it was a single value, we might need to adjust.
        # Typically, if "List Price" is a single value, it's the base price (Clear).
        
        if not (isinstance(list_price_raw, list) and len(list_price_raw) == 3):
             finish_multiplier = 1.0
             finish_norm = finish.lower() if finish else 'clear'
             if finish_norm == 'black':
                 finish_multiplier = 1.1
             elif finish_norm == 'paint':
                 finish_multiplier = 1.2
             list_price_effective *= finish_multiplier

        # 3. Divide by length to get per-foot price for profiles
        # CRITICAL FIX: Ensure length_ft is calculated correctly.
        # BE9-2513 has Length: "24' - 0"" which parse_length_to_feet handles.
        length_ft = parse_length_to_feet(length_str)
        if length_ft > EPSILON: 
            list_price_effective /= length_ft
        else:
            print(f"Warning: Profile '{part_number}' has zero or invalid length '{length_str}'. Unit price might be incorrect.")

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


def get_price_by_part(part_number, requested_qty, finish=None, current_extra_materials=None, summary=False, group=False, extra_materials_file="extra_materials.json", description=None):
    """
    Calculate price and material impact, considering finish for profiles.
    
    Args:
        part_number (str): The part number.
        requested_qty (float/int/list): The quantity requested. Can be a list for bay width parts.
        finish (str, optional): The finish type ('clear', 'black', 'paint'). Only relevant for profiles.
        current_extra_materials (dict, optional): In-memory extra materials state.
        summary (bool): If True, only return price and unit type, no material impact details.
        group (bool): If True, forces profile-like behavior for pricing.
        extra_materials_file (str): Path to the extra materials JSON file.
        description (str, optional): Description of the part, used to distinguish jamb vs sill for BE9-2513.
    
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
    
    # DEBUG LOGGING
    if unit_price is None or unit_price == 0:
        with open("pricing_debug_log.txt", "a") as log:
            log.write(f"Zero/None Price for {part_number}: UnitPrice={unit_price}, Type={unit_type}, Finish={finish}\n")
    
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
        'finish': finish, # Store the finish in material impact details
        'description': description  # Store description for reference
    }

    if is_profile_for_inventory: # This logic handles profiles and length-based items (like E2-0052 now that unit_type is 'ft')
        unit_type = "ft" # Confirm unit type for reporting
        min_purchase_length = parse_length_to_feet(length_str) or 1.0
        leftover_pieces_sim = sorted(list(part_extra_sim.get('length_pieces', [])), reverse=True)  # Sort descending for better matching
        
        # Check if this is a bay width part (for BE9-2513, check if requested_qty is a list to distinguish sill from jamb)
        # Use provided description or fall back to material_impact_details
        part_description = description or material_impact_details.get('description')
        is_bay_width_part = _is_bay_width_part(part_number, requested_qty, part_description)
        
        # Handle list of bay widths for special parts
        if isinstance(requested_qty, list) and is_bay_width_part and not summary:
            # For bay width parts, optimize across the entire list
            bay_widths = [float(q) for q in requested_qty]
            total_needed = sum(bay_widths)
            used_from_leftover = 0.0
            remaining_leftovers = leftover_pieces_sim.copy()
            matched_bays = []
            
            # Try to match leftover pieces to bay widths
            # Strategy: Process bay widths in descending order to avoid using large pieces for small requirements
            # For each bay_width, find the smallest leftover piece that is >= bay_width (closest fit)
            # This minimizes waste by using the closest fit that doesn't exceed by much
            leftover_pieces_consumed = []  # Track which leftover pieces were used
            # Sort bay widths in descending order to match largest pieces to largest requirements first
            # This prevents using a large leftover piece for a small bay when it could be used for a larger bay
            sorted_bay_widths = sorted(bay_widths, reverse=True)
            
            for bay_width in sorted_bay_widths:
                best_match_index = None
                best_match_length = None
                
                # Find the smallest leftover piece that is >= bay_width (closest fit to minimize waste)
                # This ensures we use the piece that "doesn't exceed it" by much
                for i, leftover in enumerate(remaining_leftovers):
                    if leftover >= bay_width - EPSILON:
                        # This piece is large enough
                        if best_match_index is None or leftover < best_match_length:
                            # This is the smallest suitable piece found so far (best fit - closest without exceeding much)
                            best_match_index = i
                            best_match_length = leftover
                
                if best_match_index is not None:
                    # Use the closest fitting piece
                    leftover = remaining_leftovers[best_match_index]
                    used_from_leftover += bay_width
                    remaining_after_use = leftover - bay_width
                    # Track this consumption
                    leftover_pieces_consumed.append({
                        'original_length': leftover,
                        'used_length': bay_width
                    })
                    remaining_leftovers.pop(best_match_index)
                    if remaining_after_use > EPSILON:
                        # Insert remaining piece back in sorted order
                        remaining_leftovers.append(remaining_after_use)
                        remaining_leftovers.sort(reverse=True)
                    matched_bays.append(bay_width)
                else:
                    # No suitable leftover piece found for this bay
                    matched_bays.append(bay_width)
            
            remaining_needed = total_needed - used_from_leftover
            
            if remaining_needed > EPSILON:
                # Purchase new material for remaining needs
                num_bundles_needed = math.ceil(remaining_needed / min_purchase_length)
                actual_purchased_length = num_bundles_needed * min_purchase_length
                total_price = unit_price * actual_purchased_length
                
                # Calculate leftover after using for remaining bays
                leftover_after_use = actual_purchased_length - remaining_needed
                
                # If leftover is >= any bay width, it can be reused
                if leftover_after_use > EPSILON:
                    material_impact_details['leftover_generated_qty_or_length'] = leftover_after_use
                
                material_impact_details['purchased_qty_or_length'] = actual_purchased_length
                material_impact_details['cost_incurred'] = total_price
            else:
                # All needs met from leftovers
                total_price = 0.0
                material_impact_details['purchased_qty_or_length'] = 0.0
                material_impact_details['cost_incurred'] = 0.0
            
            material_impact_details['used_from_leftover_qty_or_length'] = used_from_leftover
            material_impact_details['bay_widths_processed'] = matched_bays  # Store for debugging
            material_impact_details['is_bay_width_list'] = True  # Flag to indicate special handling needed
            # Store information about which leftover pieces were used (for proper removal in apply_material_impact)
            material_impact_details['leftover_pieces_consumed'] = leftover_pieces_consumed
            
        else:
            # Standard handling for single quantity or non-bay-width parts
            # Convert list to sum if it's a list but not a bay width part
            if isinstance(requested_qty, list):
                requested_qty = sum(requested_qty)
            
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
        # Handle bay width lists specially
        is_bay_width_list = material_impact_details.get('is_bay_width_list', False)
        leftover_pieces_consumed = material_impact_details.get('leftover_pieces_consumed', [])
        
        if is_bay_width_list and leftover_pieces_consumed:
            # Process each consumed leftover piece
            temp_leftovers = sorted(list(part_extra.get('length_pieces', [])), reverse=True)
            for consumed_info in leftover_pieces_consumed:
                original_length = consumed_info.get('original_length')
                used_length = consumed_info.get('used_length')
                
                # Find and remove the matching leftover piece
                found = False
                for i, piece_len in enumerate(temp_leftovers):
                    # Match by original length (with tolerance)
                    if abs(piece_len - original_length) < EPSILON:
                        remaining_after_use = piece_len - used_length
                        temp_leftovers.pop(i)
                        if remaining_after_use > EPSILON:
                            temp_leftovers.append(remaining_after_use)
                            temp_leftovers.sort(reverse=True)
                        found = True
                        break
                
                if not found:
                    # Fallback: find any piece >= used_length
                    for i, piece_len in enumerate(temp_leftovers):
                        if piece_len >= used_length - EPSILON:
                            remaining_after_use = piece_len - used_length
                            temp_leftovers.pop(i)
                            if remaining_after_use > EPSILON:
                                temp_leftovers.append(remaining_after_use)
                                temp_leftovers.sort(reverse=True)
                            found = True
                            break
                    
                    if not found:
                        print(f"⚠️ Warning: Could not find suitable leftover piece to consume {used_length:.4f} for {part_number} ({finish}).")
            
            part_extra['length_pieces'] = temp_leftovers
            part_extra['quantity'] = 0.0
        else:
            # Standard handling for single quantity
            used_from_leftover_qty_or_length = material_impact_details.get('used_from_leftover_qty_or_length', 0.0)
            if used_from_leftover_qty_or_length > EPSILON:
                temp_leftovers = sorted(list(part_extra.get('length_pieces', [])), reverse=True)
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
        # Handle bay width lists specially
        is_bay_width_list = material_impact_details.get('is_bay_width_list', False)
        leftover_pieces_consumed = material_impact_details.get('leftover_pieces_consumed', [])
        
        if is_bay_width_list and leftover_pieces_consumed:
            # Process each consumed leftover piece
            temp_leftovers = sorted(list(part_extra.get('length_pieces', [])), reverse=True)
            for consumed_info in leftover_pieces_consumed:
                original_length = consumed_info.get('original_length')
                used_length = consumed_info.get('used_length')
                
                # Find and remove the matching leftover piece
                found = False
                for i, piece_len in enumerate(temp_leftovers):
                    # Match by original length (with tolerance)
                    if abs(piece_len - original_length) < EPSILON:
                        remaining_after_use = piece_len - used_length
                        temp_leftovers.pop(i)
                        if remaining_after_use > EPSILON:
                            temp_leftovers.append(remaining_after_use)
                            temp_leftovers.sort(reverse=True)
                        found = True
                        break
                
                if not found:
                    # Fallback: find any piece >= used_length
                    for i, piece_len in enumerate(temp_leftovers):
                        if piece_len >= used_length - EPSILON:
                            remaining_after_use = piece_len - used_length
                            temp_leftovers.pop(i)
                            if remaining_after_use > EPSILON:
                                temp_leftovers.append(remaining_after_use)
                                temp_leftovers.sort(reverse=True)
                            found = True
                            break
                    
                    if not found:
                        print(f"⚠️ Warning (in-memory): Could not find suitable leftover piece to consume {used_length:.4f} for {part_number} ({finish}).")
            
            part_extra['length_pieces'] = temp_leftovers
            part_extra['quantity'] = 0.0
        else:
            # Standard handling for single quantity
            used_from_leftover_qty_or_length = material_impact_details.get('used_from_leftover_qty_or_length', 0.0)
            if used_from_leftover_qty_or_length > EPSILON:
                temp_leftovers = sorted(list(part_extra.get('length_pieces', [])), reverse=True)
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