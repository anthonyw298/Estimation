from utils.formulas import calculate_door_info

v= calculate_door_info([{'size': "3' X 7'", 'count': 1, 'stile': 'Narrow', 'hardware': ['Continuous Hinges']}])
print(v)

door_items = calculate_door_info(doors) if doors else []
        calculated_outputs.extend(door_items)