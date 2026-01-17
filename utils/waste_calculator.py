"""
Waste Calculator Enhancement Module
Provides visual waste percentage impact, waste breakdown by material type, and optimization suggestions.
"""
import os
import json
from typing import Dict, List, Tuple, Optional
from data.parts_data import parts_data
from utils.pricing import parse_length_to_feet, get_unit_price_by_part, load_extra_materials
from collections import Counter

def read_waste_from_excel(excel_path: str) -> Dict:
    """
    Read waste statistics directly from Excel Summary sheet.
    
    Returns:
        Dictionary with waste statistics read from Excel:
        - total_waste_cost: Residual/Waste Cost from Excel
        - total_material_cost: Discounted Total from Excel
        - overall_waste_percentage: Waste Percentage from Excel
    """
    try:
        from openpyxl import load_workbook
        
        if not os.path.exists(excel_path):
            return None
        
        wb = load_workbook(excel_path, data_only=True)
        
        # Check if Summary sheet exists
        if "Summary" not in wb.sheetnames:
            if "SUMMARY" in wb.sheetnames:
                ws = wb["SUMMARY"]
            else:
                return None
        else:
            ws = wb["Summary"]
        
        # Search for "COST OVERVIEW" to find the start position
        overview_row = None
        overview_col = None
        
        for row in range(1, min(ws.max_row + 1, 100)):  # Search first 100 rows
            for col in range(1, min(ws.max_column + 1, 10)):  # Search first 10 columns
                cell_value = ws.cell(row=row, column=col).value
                if cell_value and "COST OVERVIEW" in str(cell_value).upper():
                    overview_row = row
                    overview_col = col
                    break
            if overview_row:
                break
        
        if not overview_row:
            return None
        
        # Read values from relative positions
        # Discounted Total is at overview_row+2, overview_col+2
        # Residual/Waste Cost is at overview_row+3, overview_col+2
        # Waste Percentage is at overview_row+4, overview_col+2
        
        discounted_total_cell = ws.cell(row=overview_row + 2, column=overview_col + 2)
        waste_cost_cell = ws.cell(row=overview_row + 3, column=overview_col + 2)
        waste_percentage_cell = ws.cell(row=overview_row + 4, column=overview_col + 2)
        
        # Extract values
        total_material_cost = float(discounted_total_cell.value) if discounted_total_cell.value else 0.0
        total_waste_cost = float(waste_cost_cell.value) if waste_cost_cell.value else 0.0
        
        # Extract waste percentage (remove % sign if present)
        waste_pct_str = str(waste_percentage_cell.value) if waste_percentage_cell.value else "0"
        waste_pct_str = waste_pct_str.replace("%", "").strip()
        overall_waste_percentage = float(waste_pct_str) if waste_pct_str else 0.0
        
        return {
            "total_waste_cost": total_waste_cost,
            "total_material_cost": total_material_cost,
            "overall_waste_percentage": overall_waste_percentage
        }
    except Exception as e:
        print(f"⚠️ Error reading waste from Excel: {e}")
        import traceback
        traceback.print_exc()
        return None

def calculate_waste_statistics(project_path: str, extra_materials_path: str, excel_path: Optional[str] = None) -> Dict:
    """
    Calculate comprehensive waste statistics for a project.
    
    Returns:
        Dictionary with waste statistics including:
        - total_waste_cost: Total cost of waste materials
        - total_material_cost: Total cost of materials
        - overall_waste_percentage: Overall waste percentage
        - material_breakdown: List of waste data per material
        - suggestions: List of optimization suggestions
    """
    if not os.path.exists(project_path) or not os.path.exists(extra_materials_path):
        return {
            "total_waste_cost": 0.0,
            "total_material_cost": 0.0,
            "overall_waste_percentage": 0.0,
            "material_breakdown": [],
            "suggestions": []
        }
    
    try:
        # Load elevations data
        with open(project_path, 'r') as f:
            elevations_data = json.load(f)
        
        # Load extra materials (leftovers/waste)
        extra_materials = load_extra_materials(extra_materials_path)
        
        # Debug output
        print(f"📊 Waste Calculator: Loaded {len(elevations_data)} elevations, {len(extra_materials)} extra materials")
    except Exception as e:
        print(f"❌ Error loading waste data: {e}")
        import traceback
        traceback.print_exc()
        return {
            "total_waste_cost": 0.0,
            "total_material_cost": 0.0,
            "overall_waste_percentage": 0.0,
            "material_breakdown": [],
            "suggestions": []
        }
    
    # Calculate waste per material
    material_breakdown = []
    total_waste_cost = 0.0
    total_material_cost = 0.0
    
    # Process each material in extra_materials (these are the waste/leftovers)
    for material_key, material_data in extra_materials.items():
        if not material_data:
            continue
        
        # Extract part number and finish from key (format: "BE9-2513-clear" or "E1-0199" or "PM-1006-SS")
        parts = material_key.split('-')
        # Common finish codes
        finish_codes = ['clear', 'bronze', 'grey', 'black', 'white']
        
        if len(parts) >= 3:
            # Check if last part is a finish code
            if parts[-1].lower() in finish_codes:
                # Format: "BE9-2513-clear" -> part_number = "BE9-2513", finish = "clear"
                part_number = '-'.join(parts[:-1])
                finish = parts[-1].lower()
            else:
                # Format: "PM-1006-SS" -> part_number = "PM-1006-SS", finish = ""
                part_number = material_key
                finish = ""
        elif len(parts) == 2:
            # Could be "E1-0199" (no finish) or "BE9-clear" (finish)
            if parts[1].lower() in finish_codes:
                part_number = parts[0]
                finish = parts[1].lower()
            else:
                # No finish code, entire key is part number
                part_number = material_key
                finish = ""
        else:
            part_number = material_key
            finish = ""
        
        # Get material info
        part_info = parts_data.get(part_number, {})
        description = part_info.get('Description', part_number)
        
        # Calculate waste quantity and cost
        length_pieces = material_data.get('length_pieces', [])
        quantity = material_data.get('quantity', 0.0)
        
        # Calculate waste quantity (from leftover pieces or quantity)
        waste_qty = 0.0
        if length_pieces and len(length_pieces) > 0:
            # For profiles/gaskets, waste is in length_pieces
            length_str = part_info.get('Length', '')
            min_purchase_length = parse_length_to_feet(length_str) or 24.0
            valid_lengths = []
            for l in length_pieces:
                try:
                    l_float = float(l)
                    # Valid leftover: > 0 and < min_purchase_length
                    if 0 < l_float < min_purchase_length:
                        valid_lengths.append(l_float)
                except (TypeError, ValueError):
                    pass
            waste_qty = sum(valid_lengths)
        elif quantity and float(quantity) > 0:
            # For accessories, waste is in quantity field
            try:
                waste_qty = float(quantity)
            except (TypeError, ValueError):
                waste_qty = 0.0
        
        if waste_qty <= 0:
            continue
        
        # Calculate material cost (using unit price)
        unit_price, _ = get_unit_price_by_part(
            part_number, 
            finish=finish, 
            extra_materials_file=extra_materials_path
        )
        
        # Apply multiplier based on total cost (same logic as Excel generator)
        # Get total project cost to determine multiplier
        total_project_cost = 0.0
        for elev_key, elev in elevations_data.items():
            for output in elev.get('calculated_outputs', []):
                qty = output.get('quantity', 0)
                if isinstance(qty, list):
                    qty_val = sum([float(x) for x in qty if x])
                else:
                    qty_val = float(qty) if qty else 0.0
                price = output.get('price', 0.0)
                total_project_cost += qty_val * float(price) if price else 0.0
        
        # Use same multiplier logic as Excel generator (0.614 if < 50000, 0.572 if >= 50000)
        multiplier = 0.614 if total_project_cost < 50000 else 0.572
        waste_cost = waste_qty * unit_price * multiplier if unit_price else 0.0
        
        # Calculate total material quantity used for this part (sum across all elevations)
        # This is the quantity that was actually used (before waste)
        total_qty_used = 0.0
        for elev_key, elev in elevations_data.items():
            elev_finish = elev.get('finish', '').lower()
            for output in elev.get('calculated_outputs', []):
                output_part = output.get('part_number', '').strip()
                # Match part number (exact match or match without finish suffix)
                part_matches = False
                if output_part == part_number:
                    part_matches = True
                elif not finish and output_part.startswith(part_number):
                    # If no finish in key, try partial match (e.g., "PM-1006-SS" matches "PM-1006-SS")
                    part_matches = True
                
                if part_matches:
                    # If material key has finish, check if elevation finish matches
                    if finish and elev_finish != finish.lower():
                        continue
                    qty = output.get('quantity', 0)
                    if isinstance(qty, list):
                        total_qty_used += sum([float(x) for x in qty if x and x > 0])
                    else:
                        try:
                            total_qty_used += float(qty) if qty else 0.0
                        except (TypeError, ValueError):
                            pass
        
        # If we don't have usage data, try to estimate from purchased quantity
        # Purchased quantity = used + waste
        total_qty_purchased = total_qty_used + waste_qty
        
        # Skip if no waste and no usage
        if waste_qty <= 0 and total_qty_used <= 0:
            continue
        
        # If we have waste but no usage data, debug the matching issue
        if total_qty_used <= 0 and waste_qty > 0:
            # Try more lenient matching - check if any part_number starts with our part
            for elev_key, elev in elevations_data.items():
                for output in elev.get('calculated_outputs', []):
                    output_part = output.get('part_number', '').strip()
                    # Try exact match, or partial match (for cases like "PM-1006-SS")
                    if output_part == part_number or output_part == material_key:
                        qty = output.get('quantity', 0)
                        if isinstance(qty, list):
                            total_qty_used += sum([float(x) for x in qty if x and x > 0])
                        else:
                            try:
                                total_qty_used += float(qty) if qty else 0.0
                            except (TypeError, ValueError):
                                pass
            
            # If still no usage found, skip to avoid 100% waste calculation
            if total_qty_used <= 0:
                print(f"⚠️ Waste Calculator: Could not match usage data for {material_key} (waste: {waste_qty}, tried part_number: {part_number})")
                continue
            
            # Recalculate purchased quantity with found usage
            total_qty_purchased = total_qty_used + waste_qty
        
        # Calculate waste percentage for this material
        # Waste percentage = waste / (used + waste) * 100
        if total_qty_purchased > 0:
            waste_percentage = (waste_qty / total_qty_purchased * 100)
        else:
            waste_percentage = 0.0
        
        # Calculate material cost for used materials only (not including waste)
        # This matches the Excel calculation: waste_cost / total_discounted_price * 100
        # where total_discounted_price is the cost of USED materials only
        used_material_cost = total_qty_used * unit_price * multiplier if unit_price else 0.0
        
        material_breakdown.append({
            "part_number": part_number,
            "description": description,
            "finish": finish,
            "total_quantity": total_qty_used,
            "waste_quantity": waste_qty,
            "waste_percentage": waste_percentage,
            "waste_cost": waste_cost,
            "material_cost": used_material_cost,
            "unit": "ft" if length_pieces and len(length_pieces) > 0 else "pcs"
        })
        
        total_waste_cost += waste_cost
        total_material_cost += used_material_cost  # Only count used materials, not waste
    
    # Try to read waste statistics directly from Excel Summary sheet (more accurate)
    excel_waste_data = None
    if excel_path and os.path.exists(excel_path):
        excel_waste_data = read_waste_from_excel(excel_path)
    
    # Use Excel values if available, otherwise calculate from data
    if excel_waste_data:
        overall_waste_percentage = excel_waste_data["overall_waste_percentage"]
        total_waste_cost = excel_waste_data["total_waste_cost"]
        total_material_cost = excel_waste_data["total_material_cost"]
        print(f"✅ Waste Calculator: Using values from Excel (waste: {overall_waste_percentage:.2f}%, cost: ${total_waste_cost:.2f})")
    else:
        # Calculate overall waste percentage to match Excel report
        # Excel uses: waste_cost / total_discounted_price * 100
        # where total_discounted_price is the cost of USED materials (not including waste)
        overall_waste_percentage = (total_waste_cost / total_material_cost * 100) if total_material_cost > 0 else 0.0
        print(f"⚠️ Waste Calculator: Excel not available, calculated waste: {overall_waste_percentage:.2f}%")
    
    # Generate optimization suggestions
    suggestions = generate_optimization_suggestions(material_breakdown, overall_waste_percentage)
    
    return {
        "total_waste_cost": total_waste_cost,
        "total_material_cost": total_material_cost,
        "overall_waste_percentage": overall_waste_percentage,
        "material_breakdown": sorted(material_breakdown, key=lambda x: x['waste_cost'], reverse=True),
        "suggestions": suggestions
    }

def generate_optimization_suggestions(material_breakdown: List[Dict], overall_waste_percentage: float) -> List[str]:
    """
    Generate optimization suggestions based on waste statistics.
    """
    suggestions = []
    
    # High overall waste percentage
    if overall_waste_percentage > 15:
        suggestions.append({
            "priority": "high",
            "message": f"Overall waste percentage is {overall_waste_percentage:.1f}%, which is high. Consider consolidating orders across multiple elevations to reduce waste."
        })
    elif overall_waste_percentage > 10:
        suggestions.append({
            "priority": "medium",
            "message": f"Overall waste percentage is {overall_waste_percentage:.1f}%. There's room for optimization by better material planning."
        })
    
    # High waste by material
    high_waste_materials = [m for m in material_breakdown if m['waste_percentage'] > 20]
    if high_waste_materials:
        top_material = high_waste_materials[0]
        suggestions.append({
            "priority": "high" if top_material['waste_percentage'] > 30 else "medium",
            "message": f"{top_material['description']} has {top_material['waste_percentage']:.1f}% waste (${top_material['waste_cost']:.2f}). Consider adjusting cut lengths or combining with other projects."
        })
    
    # High cost waste materials
    high_cost_waste = [m for m in material_breakdown if m['waste_cost'] > 500]
    if high_cost_waste:
        top_cost = high_cost_waste[0]
        suggestions.append({
            "priority": "high",
            "message": f"{top_cost['description']} waste cost is ${top_cost['waste_cost']:.2f}. Explore alternative cutting strategies to minimize this waste."
        })
    
    # Multiple small waste pieces
    small_waste_count = len([m for m in material_breakdown if 0 < m['waste_quantity'] < 2 and m['waste_percentage'] > 5])
    if small_waste_count > 3:
        suggestions.append({
            "priority": "medium",
            "message": f"Multiple materials have small leftover pieces. Consider using custom bay widths/heights to better utilize full stock lengths."
        })
    
    # No high-priority suggestions
    if not suggestions:
        suggestions.append({
            "priority": "low",
            "message": "Waste levels are acceptable. Continue current optimization practices."
        })
    
    return suggestions

def get_waste_percentage_color(waste_percentage: float) -> str:
    """
    Get color code based on waste percentage (green < 10%, yellow 10-15%, red > 15%).
    """
    if waste_percentage < 10:
        return "#4CAF50"  # Green
    elif waste_percentage < 15:
        return "#FFC107"  # Yellow/Orange
    else:
        return "#F44336"  # Red

