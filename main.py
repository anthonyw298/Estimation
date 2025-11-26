import flet as ft
import json
import os
import datetime
from openpyxl import Workbook

# Assuming your utils and systems are in their respective directories
from utils.excel_generator import generate_excel_report
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter, calculate_total_door_area, calculate_glass_to_add_back, calculate_door_info

# --- Constants & Config ---
PROJECTS_DIR = ".files"
MASTER_PROJECT_LIST_FILE = os.path.join(PROJECTS_DIR, "projects_list.json")

# United Glass color palette (matching logo)
COLOR_BG = "#000000"       # Black background (matching logo)
COLOR_SURFACE = "#1A1A1A"  # Dark grey card background
COLOR_ACCENT = "#0073E6"   # Vibrant blue (matching logo U)
COLOR_TEXT = "#FFFFFF"     # White text for readability on black
COLOR_TEXT_DIM = "#B3B3B3" # Light grey for secondary text
COLOR_INPUT_BG = "#2A2A2A" # Dark grey input background
COLOR_ACCENT_LIGHT = "#D3D3D3"  # Light grey (matching logo side surfaces)

def main(page: ft.Page):
    page.title = "United Glass Estimation"
    page.theme_mode = ft.ThemeMode.DARK  # Dark mode to match black background
    page.padding = 0
    page.bgcolor = COLOR_BG
    
    # Set Arial as requested
    page.theme = ft.Theme(
        font_family="Arial",
        color_scheme=ft.ColorScheme(
            primary=COLOR_ACCENT,
            background=COLOR_BG,
            surface=COLOR_SURFACE,
        )
    )

    # Ensure project dir exists
    os.makedirs(PROJECTS_DIR, exist_ok=True)

    # --- State ---
    state = {
        "projects": [],
        "current_project": None,
        "saved_elevations": {},
        "current_doors": [],
        "selected_door_index": None,
        "system_options": ["YES 45TU FRONT SET(OG)", "Other"],
        "finish_options": ["Clear", "Black", "Paint"],
        "door_options": ['None', "3' X 7'", "3' X 8'", "3' X 9'", "6' X 7'", "6' X 8'", "6' X 9'"],
        "stile_options": ["Narrow", "Medium", "Wide"],
        "hardware_options": [
            "Continuous Hinges", "Concealed Closer", "Exit Devices", "Electric Strike", 
            "Extended Ladder Pull (B2B)", "Extended Ladder Pull (Single)", 
            "Latch Lock w/ Lever Handle", "Lever Handle"
        ]
    }

    # --- Inputs Ref (to access values easily) ---
    inputs = {}

    # --- Data Loading Functions ---
    def load_projects():
        if os.path.exists(MASTER_PROJECT_LIST_FILE):
            try:
                with open(MASTER_PROJECT_LIST_FILE, 'r') as f:
                    state["projects"] = json.load(f)
            except:
                state["projects"] = []
        else:
            state["projects"] = []

    def save_projects():
        with open(MASTER_PROJECT_LIST_FILE, 'w') as f:
            json.dump(state["projects"], f, indent=4)

    def get_project_paths(project_name):
        clean_name = project_name.replace(" ", "_").replace("/", "_")
        base = os.path.join(PROJECTS_DIR, clean_name)
        return {
            "excel": f"{base}_Report.xlsx",
            "elevations": f"{base}_Elevations.json",
            "materials": f"{base}_ExtraMaterials.json"
        }

    def load_elevations(project_name):
        paths = get_project_paths(project_name)
        if os.path.exists(paths["elevations"]):
            try:
                with open(paths["elevations"], 'r') as f:
                    state["saved_elevations"] = json.load(f)
            except:
                state["saved_elevations"] = {}
        else:
            state["saved_elevations"] = {}
            # Create files if missing
            with open(paths["elevations"], 'w') as f: json.dump({}, f)
            if not os.path.exists(paths["materials"]):
                with open(paths["materials"], 'w') as f: json.dump({}, f)

    def get_door_path(project_name, elev_type):
        if not project_name or not elev_type: return None
        p_clean = project_name.replace(" ", "_")
        e_clean = elev_type.replace(" ", "_")
        return os.path.join(PROJECTS_DIR, f"{p_clean}_{e_clean}_doors.json")

    def load_doors(elev_type):
        path = get_door_path(state["current_project"], elev_type)
        if path and os.path.exists(path):
            try:
                with open(path, 'r') as f:
                    state["current_doors"] = json.load(f)
            except:
                state["current_doors"] = []
        else:
            state["current_doors"] = []

    def save_doors(elev_type):
        path = get_door_path(state["current_project"], elev_type)
        if path:
            with open(path, 'w') as f:
                json.dump(state["current_doors"], f, indent=4)

    # --- UI Building Blocks ---
    
    def create_input_field(label, key, expand=True, value="", numeric=False, on_change=None):
        field = ft.TextField(
            label=label,
            value=value,
            expand=expand,
            bgcolor=COLOR_INPUT_BG,
            border_color=COLOR_ACCENT_LIGHT,
            text_size=14,
            color=COLOR_TEXT,
            label_style=ft.TextStyle(color=COLOR_TEXT_DIM),
            focused_border_color=COLOR_ACCENT,
            border_radius=5,
            on_change=on_change
        )
        inputs[key] = field
        return field

    def create_dropdown(label, key, options, expand=True, on_change=None):
        field = ft.Dropdown(
            label=label,
            options=[ft.dropdown.Option(o) for o in options],
            value=options[0] if options else None,
            expand=expand,
            bgcolor=COLOR_INPUT_BG,
            border_color=COLOR_ACCENT_LIGHT,
            text_size=14,
            color=COLOR_TEXT,
            label_style=ft.TextStyle(color=COLOR_TEXT_DIM),
            focused_border_color=COLOR_ACCENT,
            border_radius=5,
            on_change=on_change
        )
        inputs[key] = field
        return field

    def show_snack(msg, color="white"):
        page.snack_bar = ft.SnackBar(content=ft.Text(msg, color=COLOR_TEXT), bgcolor=COLOR_SURFACE)
        page.snack_bar.open = True
        page.update()

        # Helper to assign ref to inputs dict easily (hacky but works for keeping existing structure)
    def assign_ref(k, control):
        inputs[k] = control
        return control

    # --- Views ---

    def build_projects_view():
        load_projects()
        
        def on_project_click(e, name):
            state["current_project"] = name
            load_elevations(name)
            page.go("/workspace")

        def add_project_click(e):
            if not new_proj_name.value: return
            if new_proj_name.value in state["projects"]:
                show_snack("Project exists!", COLOR_ACCENT)
                return
            state["projects"].append(new_proj_name.value)
            save_projects()
            # Refresh view
            page.views.pop()
            page.views.append(build_projects_view())
            page.update()

        def delete_project_click(e, name):
            if name in state["projects"]:
                state["projects"].remove(name)
                save_projects()
                
                # Clean up files
                try:
                    paths = get_project_paths(name)
                    for p in paths.values():
                        if os.path.exists(p): os.remove(p)
                    
                    # Also remove any door files
                    clean_name = name.replace(" ", "_").replace("/", "_")
                    for f in os.listdir(PROJECTS_DIR):
                        if f.startswith(clean_name) and "_doors.json" in f:
                            os.remove(os.path.join(PROJECTS_DIR, f))
                except Exception as ex:
                    print(f"Error cleaning up project files: {ex}")

                # Refresh view
                page.views.pop()
                page.views.append(build_projects_view())
                page.update()
                show_snack(f"Project '{name}' deleted", "red")

        new_proj_name = create_input_field("New Project Name", "new_proj", expand=True)
        
        # Grid of projects
        project_tiles = []
        for p in state["projects"]:
            project_tiles.append(
                ft.Container(
                    content=ft.Column([
                        ft.Row([
                            ft.Container(expand=True),
                            ft.IconButton(ft.Icons.CLOSE, icon_color="red", icon_size=16, 
                                        tooltip="Delete Project", 
                                        on_click=lambda e, name=p: delete_project_click(e, name))
                        ]),
                        ft.Icon(ft.Icons.FOLDER_OPEN, size=40, color=COLOR_ACCENT),
                        ft.Text(p, size=16, weight=ft.FontWeight.BOLD, color=COLOR_TEXT, overflow=ft.TextOverflow.ELLIPSIS)
                    ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    width=160, height=160,
                    bgcolor=COLOR_SURFACE,
                    border_radius=10,
                    padding=10,
                    on_click=lambda e, name=p: on_project_click(e, name),
                    animate=ft.Animation(200, "easeOut"),
                )
            )

        # Try to load the United Glass logo
        # Try multiple possible filenames
        logo_filenames = ["R.png", "united_glass_logo.png"]
        logo_image = None
        
        for logo_filename in logo_filenames:
            logo_path = os.path.join("assets", logo_filename)
            abs_logo_path = os.path.abspath(logo_path)
            
            # Check if logo file exists and try to load it
            if os.path.exists(abs_logo_path):
                try:
                    # Try with just filename (if assets_dir is set)
                    logo_image = ft.Image(
                        src=logo_filename,
                        width=200,
                        height=200,
                        fit=ft.ImageFit.CONTAIN,
                    )
                    break  # Successfully loaded, exit loop
                except:
                    try:
                        # Try with absolute path
                        logo_image = ft.Image(
                            src=abs_logo_path,
                            width=200,
                            height=200,
                            fit=ft.ImageFit.CONTAIN,
                        )
                        break  # Successfully loaded, exit loop
                    except Exception as e:
                        print(f"Error loading {logo_filename}: {e}")
                        continue  # Try next filename
        
        # Always create a visible placeholder (will be used if image not found or as fallback)
        # Make it very visible with bright blue color
        placeholder = ft.Container(
            content=ft.Column([
                ft.Container(
                    content=ft.Text("U", size=140, weight=ft.FontWeight.BOLD, color=COLOR_ACCENT, font_family="Arial"),
                    width=160,
                    height=160,
                    alignment=ft.alignment.center,
                    border_radius=10,
                ),
                ft.Text("United Glass", size=22, weight=ft.FontWeight.BOLD, color=COLOR_ACCENT, font_family="Arial")
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=8, tight=True),
            width=200,
            height=200,
            alignment=ft.alignment.center,
            padding=10
        )
        
        # Use logo image if available, otherwise use placeholder
        # Always default to placeholder to ensure something is visible
        logo_display = placeholder
        if logo_image is not None:
            logo_display = logo_image
        
        return ft.View(
            "/",
            [
                ft.Container(
                    content=ft.Column([
                        # Header with logo on left and text on right
                        ft.Row([
                            # Logo on the left
                            ft.Container(
                                content=logo_display,
                                alignment=ft.alignment.center,
                            ),
                            # Text content on the right
                            ft.Column([
                                ft.Text("ESTIMATION TOOL", size=40, weight=ft.FontWeight.BOLD, color=COLOR_ACCENT, font_family="Arial"),
                                ft.Text("Select or create a project to begin", size=18, color=COLOR_TEXT_DIM),
                            ], spacing=8, alignment=ft.MainAxisAlignment.CENTER, expand=True),
                        ], spacing=25, alignment=ft.MainAxisAlignment.CENTER),
                        ft.Divider(color="transparent", height=10),
                        ft.Row([
                            new_proj_name,
                            ft.IconButton(ft.Icons.ADD_CIRCLE, icon_color=COLOR_ACCENT, icon_size=40, on_click=add_project_click, tooltip="Create Project")
                        ]),
                        ft.Divider(color="transparent", height=20),
                        ft.Row(project_tiles, wrap=True, spacing=20, run_spacing=20)
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=40,
                    alignment=ft.alignment.top_center
                )
            ],
            bgcolor=COLOR_BG,
            padding=0
        )

    def build_workspace_view():
        # --- Event Handlers ---
        def update_bay_visibility(e):
            # Guard against unmounted inputs
            if not inputs.get("bays_wide") or not inputs["bays_wide"].parent:
                return

            is_yes45 = inputs["system"].value == "YES 45TU FRONT SET(OG)"
            inputs["bays_wide"].parent.visible = is_yes45
            inputs["bays_tall"].parent.visible = is_yes45
            if inputs.get("custom_w_container"): inputs["custom_w_container"].visible = is_yes45
            if inputs.get("custom_h_container"): inputs["custom_h_container"].visible = is_yes45
            
            if is_yes45:
                update_dynamic_bay_inputs(None)
            
            if page.views: 
                page.update()

        def update_dynamic_bay_inputs(e):
            # Handler to regenerate dynamic inputs when bays_wide/tall change
            try:
                bw = int(inputs["bays_wide"].value) if inputs["bays_wide"].value else 0
                bh = int(inputs["bays_tall"].value) if inputs["bays_tall"].value else 0
            except:
                bw, bh = 0, 0
            
            # --- Dynamic Widths ---
            current_w_fields = inputs.get("dynamic_w_fields", [])
            # Store current values to persist if possible
            current_w_vals = [f.value for f in current_w_fields]
            
            new_w_fields = []
            inputs["custom_w_col"].controls.clear()
            
            if bw > 0:
                inputs["custom_w_col"].controls.append(ft.Text("Custom Bay Widths (leave blank to auto-fill)", size=12, color=COLOR_TEXT_DIM))
                for i in range(bw):
                    val = current_w_vals[i] if i < len(current_w_vals) else ""
                    field = create_input_field(f"Bay {i+1} Width", f"bay_w_{i}", expand=True, value=val)
                    new_w_fields.append(field)
                
                # Group fields in rows of 4 for cleaner layout
                rows = []
                for i in range(0, len(new_w_fields), 4):
                    chunk = new_w_fields[i:i+4]
                    rows.append(ft.Row(chunk))
                
                inputs["custom_w_col"].controls.extend(rows)
                
                # Add Auto-Fill Button
                def auto_fill_w(e):
                    try:
                        total_w = float(inputs["width"].value)
                    except:
                        show_snack("Please set valid Opening Width first", "red")
                        return

                    filled_sum = 0.0
                    blank_count = 0
                    
                    for f in new_w_fields:
                        try:
                            v = float(f.value)
                            filled_sum += v
                        except:
                            blank_count += 1
                    
                    if filled_sum > total_w:
                         show_snack(f"Error: Filled widths ({filled_sum:.2f}) exceed total ({total_w:.2f})", "red")
                         return

                    if blank_count > 0:
                        remaining = total_w - filled_sum
                        share = remaining / blank_count
                        for f in new_w_fields:
                            if not f.value:
                                f.value = f"{share:.4f}"
                        page.update()
                        
                        # Validate again after filling
                        filled_sum = 0.0
                        for f in new_w_fields:
                            try:
                                v = float(f.value)
                                filled_sum += v
                            except: pass
                        
                        if abs(filled_sum - total_w) > 0.01:
                             show_snack(f"Warning: Sum ({filled_sum:.2f}) does not match Total ({total_w:.2f})", "orange")
                        else:
                             show_snack("Auto-fill complete.", "green")
                    else:
                        if abs(filled_sum - total_w) > 0.01:
                             show_snack(f"Warning: Sum ({filled_sum:.2f}) does not match Total ({total_w:.2f})", "orange")
                        else:
                             show_snack("All fields filled.", "green")

                inputs["custom_w_col"].controls.append(
                    ft.ElevatedButton("Auto-Fill Remaining Widths", on_click=auto_fill_w, bgcolor=COLOR_ACCENT, color="white")
                )

            inputs["dynamic_w_fields"] = new_w_fields

            # --- Dynamic Heights ---
            current_h_fields = inputs.get("dynamic_h_fields", [])
            current_h_vals = [f.value for f in current_h_fields]
            
            new_h_fields = []
            inputs["custom_h_col"].controls.clear()
            
            if bh > 0:
                inputs["custom_h_col"].controls.append(ft.Text("Custom Bay Heights (leave blank to auto-fill)", size=12, color=COLOR_TEXT_DIM))
                for i in range(bh):
                    val = current_h_vals[i] if i < len(current_h_vals) else ""
                    field = create_input_field(f"Bay {i+1} Height", f"bay_h_{i}", expand=True, value=val)
                    new_h_fields.append(field)
                
                rows = []
                for i in range(0, len(new_h_fields), 4):
                    chunk = new_h_fields[i:i+4]
                    rows.append(ft.Row(chunk))
                
                inputs["custom_h_col"].controls.extend(rows)

                # Add Auto-Fill Button
                def auto_fill_h(e):
                    try:
                        total_h = float(inputs["height"].value)
                    except:
                        show_snack("Please set valid Opening Height first", "red")
                        return

                    filled_sum = 0.0
                    blank_count = 0
                    
                    for f in new_h_fields:
                        try:
                            v = float(f.value)
                            filled_sum += v
                        except:
                            blank_count += 1
                    
                    if filled_sum > total_h:
                         show_snack(f"Error: Filled heights ({filled_sum:.2f}) exceed total ({total_h:.2f})", "red")
                         return

                    if blank_count > 0:
                        remaining = total_h - filled_sum
                        share = remaining / blank_count
                        for f in new_h_fields:
                            if not f.value:
                                f.value = f"{share:.4f}"
                        page.update()
                        
                        # Validate again after filling
                        filled_sum = 0.0
                        for f in new_h_fields:
                            try:
                                v = float(f.value)
                                filled_sum += v
                            except: pass
                        
                        if abs(filled_sum - total_h) > 0.01:
                             show_snack(f"Warning: Sum ({filled_sum:.2f}) does not match Total ({total_h:.2f})", "orange")
                        else:
                             show_snack("Auto-fill complete.", "green")
                    else:
                        if abs(filled_sum - total_h) > 0.01:
                             show_snack(f"Warning: Sum ({filled_sum:.2f}) does not match Total ({total_h:.2f})", "orange")
                        else:
                             show_snack("All fields filled.", "green")

                inputs["custom_h_col"].controls.append(
                     ft.ElevatedButton("Auto-Fill Remaining Heights", on_click=auto_fill_h, bgcolor=COLOR_ACCENT, color="white")
                )

            inputs["dynamic_h_fields"] = new_h_fields
            if page.views:
                page.update()

        def on_elevation_load(e):
            elev_name = inputs["saved_elev"].value
            if not elev_name or elev_name == "New Elevation": 
                # If "New Elevation" selected or cleared, reset to Create mode
                clear_workspace()
                inputs["save_btn"].text = "CREATE ELEVATION"
                inputs["saved_elev"].value = None # Reset dropdown selection visual if desired or keep "New Elevation"
                page.update()
                return
            
            if elev_name not in state["saved_elevations"]:
                return

            # Switch to Update mode
            inputs["save_btn"].text = "UPDATE ELEVATION"
            
            data = state["saved_elevations"][elev_name]
            inputs["system"].value = data.get("system", state["system_options"][0])
            inputs["finish"].value = data.get("finish", state["finish_options"][0])
            inputs["type"].value = elev_name
            inputs["count"].value = str(data.get("total_count", ""))
            inputs["width"].value = str(data.get("opening_width_inches", ""))
            inputs["height"].value = str(data.get("opening_height_inches", ""))
            
            inputs["bays_wide"].value = str(data.get("bays_wide", ""))
            inputs["bays_tall"].value = str(data.get("bays_tall", ""))
            
            # Load legacy CSV if exists, else populate dynamic inputs?
            # Ideally we reconstruct dynamic inputs from stored list
            # But stored data uses list of floats.
            
            update_bay_visibility(None) # This sets visibility and triggers update_dynamic_bay_inputs
            
            # Populate dynamic fields from data
            saved_w = data.get("custom_bay_widths", [])
            saved_h = data.get("custom_bay_heights", [])
            
            # Wait for UI update cycle? No, direct update.
            # update_dynamic_bay_inputs recreated the fields. We just fill them.
            w_fields = inputs.get("dynamic_w_fields", [])
            for i, val in enumerate(saved_w):
                if i < len(w_fields): w_fields[i].value = str(val)

            h_fields = inputs.get("dynamic_h_fields", [])
            for i, val in enumerate(saved_h):
                if i < len(h_fields): h_fields[i].value = str(val)

            load_doors(elev_name)
            render_doors()
            page.update()

        def clear_workspace():
            inputs["type"].value = ""
            inputs["count"].value = ""
            inputs["width"].value = ""
            inputs["height"].value = ""
            inputs["bays_wide"].value = ""
            inputs["bays_tall"].value = ""
            inputs["save_btn"].text = "CREATE ELEVATION"
            # Clear dynamic fields
            inputs["custom_w_col"].controls.clear()
            inputs["custom_h_col"].controls.clear()
            inputs["dynamic_w_fields"] = []
            inputs["dynamic_h_fields"] = []
            
            state["current_doors"] = []
            render_doors()
            page.update()

        def render_doors():
            door_list_col.controls.clear()
            for i, door in enumerate(state["current_doors"]):
                hw_txt = ", ".join([k for k,v in door["hardware"].items() if v])
                door_list_col.controls.append(
                    ft.Container(
                        content=ft.Row([
                            ft.Column([
                                ft.Text(f"Door {i+1}", weight="bold", color=COLOR_TEXT),
                                ft.Text(f"{door['size']} | {door['stile']} Stile | Qty: {door['count']}", size=12, color=COLOR_TEXT_DIM),
                                ft.Text(f"HW: {hw_txt}", size=10, color=COLOR_TEXT_DIM, italic=True) if hw_txt else ft.Container()
                            ], expand=True),
                            ft.IconButton(ft.Icons.EDIT, icon_color="blue", on_click=lambda e, idx=i: edit_door(idx)),
                            ft.IconButton(ft.Icons.DELETE, icon_color="red", on_click=lambda e, idx=i: delete_door(idx)),
                        ]),
                        bgcolor=COLOR_INPUT_BG, padding=10, border_radius=5, margin=ft.margin.only(bottom=5)
                    )
                )
            page.update()

        def edit_door(idx):
            d = state["current_doors"][idx]
            state["selected_door_index"] = idx
            inputs["door_size"].value = d['size']
            inputs["door_count"].value = str(d['count'])
            inputs["door_stile"].value = d['stile']
            for k, cb in hardware_cbs.items():
                cb.value = d['hardware'].get(k, False)
            page.update()

        def delete_door(idx):
            is_existing_elevation = inputs["saved_elev"].value is not None and inputs["saved_elev"].value != "" and inputs["saved_elev"].value != "New Elevation"
            
            state["current_doors"].pop(idx)
            save_doors_action()
            render_doors()
            
            # Auto-update elevation if loaded
            if is_existing_elevation:
                try:
                    save_elevation_action(None)
                except Exception as e:
                    print(f"Auto-update failed after door delete: {e}")
            
            page.update()

        def save_doors_action():
            # Saves current doors to file associated with current elevation input name
            elev_name = inputs["type"].value
            if elev_name:
                save_doors(elev_name)
            else:
                print("Warning: save_doors_action called but no elevation name in type input.")

        def modify_door(action):
            # Auto-save elevation if updating a door on an existing elevation
            is_existing_elevation = inputs["saved_elev"].value is not None and inputs["saved_elev"].value != ""
            
            try:
                count = int(inputs["door_count"].value)
            except:
                show_snack("Invalid door count")
                return
                
            new_door = {
                "size": inputs["door_size"].value,
                "count": count,
                "stile": inputs["door_stile"].value,
                "hardware": {k: v.value for k,v in hardware_cbs.items()}
            }
            
            if action == "add":
                state["current_doors"].append(new_door)
            elif action == "update" and state["selected_door_index"] is not None:
                state["current_doors"][state["selected_door_index"]] = new_door
            
            render_doors()
            save_doors_action()
            
            # Clear door inputs
            inputs["door_count"].value = ""
            for cb in hardware_cbs.values(): cb.value = False
            state["selected_door_index"] = None
            
            # Auto-update elevation if loaded
            # Note: We trigger this AFTER the door list is updated in state["current_doors"]
            if is_existing_elevation:
                try:
                    save_elevation_action(None)
                except Exception as e:
                    print(f"Auto-update failed: {e}")
                    
            page.update()

        def save_elevation_action(e):
            is_update = inputs["save_btn"].text == "UPDATE ELEVATION"
            try:
                elev = inputs["type"].value.strip()
                if not elev: raise ValueError("Elevation Name Required")
                
                total = int(inputs["count"].value)
                w = float(inputs["width"].value)
                h = float(inputs["height"].value)
                
                # Simple calculations for preview/saving
                sqft = calculate_rectangle_area(w/12, h/12)
                perim = calculate_perimeter(w/12, h/12)
                
                data = {
                    "system": inputs["system"].value,
                    "finish": inputs["finish"].value,
                    "total_count": total,
                    "opening_width_inches": w,
                    "opening_height_inches": h,
                    "sqft_per_type": sqft,
                    "total_sqft": sqft * total,
                    "perimeter_ft": perim,
                    "total_perimeter_ft": perim * total
                }

                if data["system"] == "YES 45TU FRONT SET(OG)":
                    bw = int(inputs["bays_wide"].value)
                    bh = int(inputs["bays_tall"].value)
                    
                    # Helper to parse CSV (Legacy / Unused now but kept for safety if needed)
                    def parse_csv(txt, total_len, count):
                        try:
                            parts = [float(x) for x in txt.split(',') if x.strip()]
                            if len(parts) > count: return [] # Fail silently or handle
                            # fill remainder
                            rem_c = count - len(parts)
                            if rem_c > 0:
                                rem_l = (total_len - sum(parts)) / rem_c
                                parts += [rem_l] * rem_c
                            return parts
                        except: return []

                    data["bays_wide"] = bw
                    data["bays_tall"] = bh
                    
                    # Gather values from dynamic fields and validate
                    # If user partially fills, auto-fill the rest before validation
                    w_vals = []
                    dynamic_w = inputs.get("dynamic_w_fields", [])
                    if dynamic_w:
                        filled_w_sum = 0.0
                        filled_count = 0
                        for f in dynamic_w:
                            try:
                                if f.value:
                                    v = float(f.value)
                                    filled_w_sum += v
                                    filled_count += 1
                            except: pass
                        
                        if filled_count > 0: # Custom mode active
                            blank_count = len(dynamic_w) - filled_count
                            if blank_count > 0:
                                remaining = w - filled_w_sum
                                if remaining < 0:
                                     raise ValueError(f"Custom widths sum ({filled_w_sum:.2f}) exceeds Opening Width ({w:.2f})")
                                share = remaining / blank_count
                                # Fill values in memory list
                                for f in dynamic_w:
                                    try:
                                        val = float(f.value) if f.value else share
                                        w_vals.append(val)
                                    except:
                                        w_vals.append(share)
                            else:
                                # All filled, just collect
                                for f in dynamic_w:
                                    try: w_vals.append(float(f.value))
                                    except: pass

                    h_vals = []
                    dynamic_h = inputs.get("dynamic_h_fields", [])
                    if dynamic_h:
                        filled_h_sum = 0.0
                        filled_count = 0
                        for f in dynamic_h:
                            try:
                                if f.value:
                                    v = float(f.value)
                                    filled_h_sum += v
                                    filled_count += 1
                            except: pass
                        
                        if filled_count > 0:
                            blank_count = len(dynamic_h) - filled_count
                            if blank_count > 0:
                                remaining = h - filled_h_sum
                                if remaining < 0:
                                     raise ValueError(f"Custom heights sum ({filled_h_sum:.2f}) exceeds Opening Height ({h:.2f})")
                                share = remaining / blank_count
                                for f in dynamic_h:
                                    try:
                                        val = float(f.value) if f.value else share
                                        h_vals.append(val)
                                    except:
                                        h_vals.append(share)
                            else:
                                for f in dynamic_h:
                                    try: h_vals.append(float(f.value))
                                    except: pass

                    # Validate totals if custom values exist (w_vals will be populated if custom mode was active)
                    if w_vals and abs(sum(w_vals) - w) > 0.01:
                        raise ValueError(f"Custom widths sum ({sum(w_vals):.2f}) does not match Opening Width ({w:.2f})")
                    if h_vals and abs(sum(h_vals) - h) > 0.01:
                        raise ValueError(f"Custom heights sum ({sum(h_vals):.2f}) does not match Opening Height ({h:.2f})")

                    data["custom_bay_widths"] = w_vals
                    data["custom_bay_heights"] = h_vals
                    
                    data["calculated_outputs"] = calculate_yes45tu_quantities(
                        bw, bh, total, w, h, state["current_doors"]
                    )
                else:
                    data["calculated_outputs"] = []

                # Append formatted door items to calculated_outputs so they are persisted in state and file
                if state["current_doors"]:
                    door_items = calculate_door_info(state["current_doors"], finish=data["finish"])
                    data["calculated_outputs"].extend(door_items)

                state["saved_elevations"][elev] = data
                
                # Save to file
                paths = get_project_paths(state["current_project"])
                with open(paths["elevations"], 'w') as f:
                    json.dump(state["saved_elevations"], f, indent=4)
                
                # Trigger Excel Gen (silent or with notif)
                generate_excel_report(
                    excel_path=paths["excel"], 
                    elevations_json_path=paths["elevations"], 
                    extra_materials_json_path=paths["materials"],
                    system_input=data["system"], 
                    finish_input=data["finish"], 
                    elevation_type=elev, 
                    total_count=total,
                    bays_wide=data.get("bays_wide", 0), 
                    bays_tall=data.get("bays_tall", 0),
                    opening_width=w, 
                    opening_height=h, 
                    sqft_per_type=sqft, 
                    total_sqft=data["total_sqft"], 
                    perimeter_ft=perim, 
                    total_perimeter_ft=data["total_perimeter_ft"],
                    calculated_outputs=data["calculated_outputs"], 
                    doors=state["current_doors"],
                    custom_bay_widths=data.get("custom_bay_widths", []), 
                    custom_bay_heights=data.get("custom_bay_heights", [])
                )
                
                # Update dropdown
                new_opts = sorted(state["saved_elevations"].keys())
                inputs["saved_elev"].options = [ft.dropdown.Option("New Elevation")] + [ft.dropdown.Option(x) for x in new_opts]
                
                if is_update:
                    inputs["saved_elev"].value = elev
                    show_snack("Elevation Updated Successfully", "green")
                else:
                    # Clear Inputs for new creation
                    clear_workspace()
                    inputs["saved_elev"].value = None # Reset dropdown to allow loading others
                    show_snack("Elevation Created Successfully", "green")
                
                page.update()

            except Exception as ex:
                show_snack(f"Error: {str(ex)}", "red")

        def delete_elevation_action(e):
            elev = inputs["saved_elev"].value
            if elev in state["saved_elevations"]:
                del state["saved_elevations"][elev]
                # Save removal
                paths = get_project_paths(state["current_project"])
                with open(paths["elevations"], 'w') as f:
                    json.dump(state["saved_elevations"], f, indent=4)
                
                # Remove door file
                dp = get_door_path(state["current_project"], elev)
                if os.path.exists(dp): os.remove(dp)
                
                # Regenerate the report to reflect deletion
                try:
                    generate_excel_report(
                        excel_path=paths["excel"], 
                        elevations_json_path=paths["elevations"], 
                        extra_materials_json_path=paths["materials"],
                        system_input="", finish_input="", elevation_type="", total_count=0,
                        bays_wide=0, bays_tall=0, opening_width=0, opening_height=0,
                        sqft_per_type=0, total_sqft=0, perimeter_ft=0, total_perimeter_ft=0,
                        calculated_outputs=[], doors=None, 
                        delete_elevation_type=elev # Explicitly pass deleted name for cleanup if needed inside generator
                    )
                except Exception as ex:
                    print(f"Error regenerating report after delete: {ex}")

                # Update UI
                new_opts = sorted(state["saved_elevations"].keys())
                inputs["saved_elev"].options = [ft.dropdown.Option("New Elevation")] + [ft.dropdown.Option(x) for x in new_opts]
                inputs["saved_elev"].value = None
                # Force switch to CREATE mode after deletion by clearing inputs
                clear_workspace()
                # Since clear_workspace resets the button text, we just need to confirm
                show_snack("Elevation Deleted", "red")
                page.update()

        def gen_full_report(e):
            try:
                paths = get_project_paths(state["current_project"])
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                out = os.path.join("reports", f"{state['current_project']}_{ts}.xlsx")
                os.makedirs("reports", exist_ok=True)
                
                generate_excel_report(
                    out, paths["elevations"], paths["materials"],
                    "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, [], None, None, mode="export_all"
                )
                show_snack(f"Full Report Generated: {out}", "green")
            except Exception as ex:
                show_snack(f"Report Error: {ex}", "red")

        # --- UI Structure ---
        
        # Header
        header = ft.Row([
            ft.IconButton(ft.Icons.ARROW_BACK, on_click=lambda e: page.go("/"), icon_color=COLOR_TEXT),
            ft.Text(state["current_project"].upper(), size=20, weight="bold", color=COLOR_TEXT),
            ft.Container(expand=True),
            ft.ElevatedButton("GENERATE REPORT", bgcolor=COLOR_ACCENT, color="white", on_click=gen_full_report)
        ], height=60, alignment=ft.MainAxisAlignment.START)

        # Left Col: Elevation Form
        elev_options = sorted(state["saved_elevations"].keys())
        # Add "New Elevation" option at the top
        dd_options = ["New Elevation"] + elev_options
        saved_dd = create_dropdown("Load Elevation", "saved_elev", dd_options, on_change=on_elevation_load)
        
        # Form Group
        form_col = ft.Column([
            ft.Text("ELEVATION DETAILS", size=14, weight="bold", color=COLOR_ACCENT),
            ft.Row([create_dropdown("System", "system", state["system_options"], on_change=update_bay_visibility), 
                   create_dropdown("Finish", "finish", state["finish_options"])]),
            ft.Row([create_input_field("Elevation Type (Name)", "type"), create_input_field("Quantity", "count")]),
            
            # Dimensions
            ft.Container(content=ft.Column([
                ft.Text("DIMENSIONS", size=12, weight="bold", color=COLOR_TEXT_DIM),
                ft.Row([create_input_field("Opening Width (\")", "width"), create_input_field("Opening Height (\")", "height")]),
            ]), margin=ft.margin.only(top=10)),

            # Bays (Hidden by default if not Yes45)
            ft.Container(content=ft.Column([
                ft.Text("BAY CONFIGURATION", size=12, weight="bold", color=COLOR_TEXT_DIM),
                ft.Row([
                    create_input_field("Bays Wide", "bays_wide", numeric=True, on_change=update_dynamic_bay_inputs), 
                    create_input_field("Bays Tall", "bays_tall", numeric=True, on_change=update_dynamic_bay_inputs)
                ]),
            ]), margin=ft.margin.only(top=10)),
            
            # Containers for dynamic inputs
            assign_ref("custom_w_container", ft.Container(
                content=inputs.setdefault("custom_w_col", ft.Column([], spacing=10)),
                visible=False
            )),
            assign_ref("custom_h_container", ft.Container(
                content=inputs.setdefault("custom_h_col", ft.Column([], spacing=10)),
                visible=False
            )),

            ft.Container(
                content=ft.Row([
                    assign_ref("save_btn", ft.ElevatedButton("CREATE ELEVATION", bgcolor=COLOR_ACCENT, color="white", on_click=save_elevation_action, expand=True, height=50)),
                    ft.IconButton(ft.Icons.DELETE_FOREVER, icon_color="red", tooltip="Delete Elevation", on_click=delete_elevation_action)
                ]),
                margin=ft.margin.only(top=20)
            )
        ], scroll=ft.ScrollMode.AUTO, expand=True)

        # Right Col: Door Manager
        hardware_cbs = {opt: ft.Checkbox(label=opt, fill_color=COLOR_ACCENT) for opt in state["hardware_options"]}
        
        door_list_col = ft.Column(scroll=ft.ScrollMode.AUTO, spacing=5, height=200)

        door_col = ft.Column([
            ft.Text("DOOR MANAGER", size=14, weight="bold", color=COLOR_ACCENT),
            ft.Row([create_dropdown("Size", "door_size", state["door_options"]), create_input_field("Count (Per Elevation)", "door_count")]),
            create_dropdown("Style", "door_stile", state["stile_options"]),
            ft.Text("Hardware:", size=12, color=COLOR_TEXT_DIM),
            ft.Column([cb for cb in hardware_cbs.values()], spacing=0),
            ft.Row([
                ft.ElevatedButton("ADD", bgcolor=COLOR_ACCENT, color="white", on_click=lambda e: modify_door("add"), expand=True),
                ft.ElevatedButton("UPDATE", bgcolor=COLOR_ACCENT, color="white", on_click=lambda e: modify_door("update"), expand=True),
            ]),
            ft.Divider(color=COLOR_SURFACE),
            door_list_col
        ], scroll=ft.ScrollMode.AUTO, expand=True)

        # Main Layout
        main_content = ft.Row([
            ft.Container(
                content=ft.Column([saved_dd, ft.Divider(color="transparent"), form_col]),
                expand=2, bgcolor=COLOR_SURFACE, border_radius=10, padding=20, margin=10
            ),
            ft.Container(
                content=door_col,
                expand=1, bgcolor=COLOR_SURFACE, border_radius=10, padding=20, margin=10
            )
        ], expand=True)

        # Initial visibility check
        # The view is returned but not yet added to page.views in route_change,
        # so page.update() in update_bay_visibility fails if called directly.
        # Instead, we just set the initial visibility state of the controls.
        
        is_yes45_init = inputs["system"].value == "YES 45TU FRONT SET(OG)"
        
        # Safely set initial visibility on containers if they exist
        if inputs.get("bays_wide") and inputs["bays_wide"].parent:
             inputs["bays_wide"].parent.visible = is_yes45_init
        if inputs.get("bays_tall") and inputs["bays_tall"].parent:
             inputs["bays_tall"].parent.visible = is_yes45_init
             
        if inputs.get("custom_w_container"): inputs["custom_w_container"].visible = is_yes45_init
        if inputs.get("custom_h_container"): inputs["custom_h_container"].visible = is_yes45_init
        
        if is_yes45_init:
            # Manually populate dynamic inputs without triggering page.update
            # Copy logic from update_dynamic_bay_inputs but remove page.update()
            try:
                bw = int(inputs["bays_wide"].value) if inputs["bays_wide"].value else 0
                bh = int(inputs["bays_tall"].value) if inputs["bays_tall"].value else 0
            except:
                bw, bh = 0, 0
            
            # Helper to reuse creation logic without update
            def populate_dynamic_fields_init(count, prefix, col_key, store_key):
                new_fields = []
                inputs[col_key].controls.clear()
                if count > 0:
                    inputs[col_key].controls.append(ft.Text(f"Custom Bay {'Widths' if 'w' in prefix else 'Heights'} (leave blank to auto-fill)", size=12, color=COLOR_TEXT_DIM))
                    saved_vals = [f.value for f in inputs.get(store_key, [])]
                    for i in range(count):
                        val = saved_vals[i] if i < len(saved_vals) else ""
                        field = create_input_field(f"Bay {i+1} {'Width' if 'w' in prefix else 'Height'}", f"{prefix}_{i}", expand=True, value=val)
                        new_fields.append(field)
                    rows = []
                    for i in range(0, len(new_fields), 4):
                        rows.append(ft.Row(new_fields[i:i+4]))
                    inputs[col_key].controls.extend(rows)
                    # Add button (dummy handler for init)
                    # We need the real handlers here for them to work immediately
                    # Accessing them from outer scope is tricky if they rely on closures inside update_dynamic_bay_inputs
                    # For initialization, we can just call update_dynamic_bay_inputs(None) safely because 
                    # page.update is now guarded inside it.
                    pass
                inputs[store_key] = new_fields

            # Actually, simpler: just let the update_bay_visibility run AFTER the view is mounted.
            # But Flet doesn't have a standard "on_mount" for views easily accessible here without refs.
            # Best approach: Defer the update until the view is returned.
            pass 
            
            # Call logic to generate fields
            # We can call the existing function because we added the page.views check inside it
            update_dynamic_bay_inputs(None)

        return ft.View(
            "/workspace",
            [
                ft.Container(content=header, padding=ft.padding.symmetric(horizontal=20), bgcolor=COLOR_BG),
                main_content
            ],
            bgcolor=COLOR_BG,
            padding=0
        )

    def route_change(e):
        page.views.clear()
        page.views.append(build_projects_view())
        if page.route == "/workspace" and state["current_project"]:
            ws_view = build_workspace_view()
            page.views.append(ws_view)
            # Now that controls are in the view list (technically), try updating visibility?
            # Actually, safer to trigger the event manually or call the logic without page.update() inside build
            
            # Let's just trigger the update logic safely
            # We need to access the inputs from the build_workspace_view closure.
            # But inputs is in the outer scope 'build_workspace_view'. 
            # We can't easily access it from here unless we return it or it's global. 
            # 'inputs' is defined in 'main' scope, so it IS accessible.
            
            # Call update_bay_visibility but catch errors or modify it to handle 'page not ready'
            pass 
            
        page.update()
        
        # Post-update check for workspace
        if page.route == "/workspace" and state["current_project"]:
             # Now controls are mounted
             # Trigger visibility update logic
             is_yes45 = inputs["system"].value == "YES 45TU FRONT SET(OG)"
             inputs["bays_wide"].parent.visible = is_yes45
             inputs["bays_tall"].parent.visible = is_yes45
             inputs["custom_w_container"].visible = is_yes45
             inputs["custom_h_container"].visible = is_yes45
             # We need to trigger dynamic input generation if needed
             # We can't call update_dynamic_bay_inputs directly easily because it's nested.
             # But wait, inputs are global-ish in this function scope.
             # The issue is 'update_dynamic_bay_inputs' is defined INSIDE build_workspace_view.
             
             # FIX: Move the initial call inside build_workspace_view but wrap page.update in a check.
             pass

    def route_change(e):
        page.views.clear()
        page.views.append(build_projects_view())
        if page.route == "/workspace" and state["current_project"]:
            page.views.append(build_workspace_view())
        page.update()

    def view_pop(e):
        page.views.pop()
        top_view = page.views[-1]
        page.go(top_view.route)

    page.on_route_change = route_change
    page.on_view_pop = view_pop
    page.go(page.route)

if __name__ == "__main__":
    ft.app(target=main, assets_dir="assets")
