def get_unit_price_by_part(part_number):
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