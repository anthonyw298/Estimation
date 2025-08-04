def add_door(self):
    if not self.current_project_name:
        self.update_status("Error: Please select a project first.", self.error_color)
        return

    try:
        door_size = self.vars['door_size'].get()
        if door_size == 'None':
            self.update_status("Error: Cannot add 'None' as a door.", self.error_color)
            return

        door_count_str = self.vars['door_count'].get()
        if not door_count_str:
            self.update_status("Error: Please enter a number of doors.", self.error_color)
            return
        door_count = int(door_count_str)

        stile_style = self.vars['stile'].get()
        hardware = [opt for opt, var in self.hardware_vars.items() if var.get()]

        new_door = {
            'size': door_size,
            'count': door_count,
            'stile': stile_style,
            'hardware': hardware
        }

        self.current_elevation_doors.append(new_door)
        self.update_door_listbox()
        self.clear_door_form()

        # Save to JSON as before
        elev_type = self.vars['elevation_type'].get().strip()
        if elev_type and elev_type in self.saved_elevations:
            self.saved_elevations[elev_type]['doors'] = self.current_elevation_doors
            with open(self.current_elevations_json_path, 'w') as f:
                json.dump(self.saved_elevations, f, indent=4)

        # ALSO save to Excel here:
        self.save_doors_to_excel()

        self.update_status("Door added successfully.", self.success_color)

    except ValueError:
        self.update_status("Error: Number of doors must be an integer.", self.error_color)


def add_door(self):
    if not self.current_project_name:
        self.update_status("Error: Please select a project first.", self.error_color)
        return
    
    try:
        door_size = self.vars['door_size'].get()
        if door_size == 'None':
            self.update_status("Error: Cannot add 'None' as a door.", self.error_color)
            return
        
        door_count_str = self.vars['door_count'].get()
        if not door_count_str:
            self.update_status("Error: Please enter a number of doors.", self.error_color)
            return
        door_count = int(door_count_str)
        
        stile_style = self.vars['stile'].get()
        hardware = [opt for opt, var in self.hardware_vars.items() if var.get()]
        
        new_door = {
            'size': door_size,
            'count': door_count,
            'stile': stile_style,
            'hardware': hardware
        }
        
        self.current_elevation_doors.append(new_door)
        self.update_door_listbox()
        self.clear_door_form()
        self.update_status("Door added successfully.", self.success_color)

    except ValueError:
        self.update_status("Error: Number of doors must be an integer.", self.error_color)