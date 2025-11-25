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

# Netflix-inspired palette
COLOR_BG = "#141414"       # Deep dark background
COLOR_SURFACE = "#1f1f1f"  # Card background
COLOR_ACCENT = "#E50914"   # Netflix Red
COLOR_TEXT = "#FFFFFF"
COLOR_TEXT_DIM = "#B3B3B3"
COLOR_INPUT_BG = "#333333"

def main(page: ft.Page):
    page.title = "United Glass Estimation"
    page.theme_mode = ft.ThemeMode.DARK
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
    
    def create_input_field(label, key, expand=True, value="", numeric=False):
        field = ft.TextField(
            label=label,
            value=value,
            expand=expand,
            bgcolor=COLOR_INPUT_BG,
            border_color="transparent",
            text_size=14,
            color="white",
            label_style=ft.TextStyle(color=COLOR_TEXT_DIM),
            focused_border_color=COLOR_ACCENT,
            border_radius=5
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
            border_color="transparent",
            text_size=14,
            color="white",
            label_style=ft.TextStyle(color=COLOR_TEXT_DIM),
            focused_border_color=COLOR_ACCENT,
            border_radius=5,
            on_change=on_change
        )
        inputs[key] = field
        return field

    def show_snack(msg, color="white"):
        page.snack_bar = ft.SnackBar(content=ft.Text(msg, color="white"), bgcolor=COLOR_SURFACE)
        page.snack_bar.open = True
        page.update()

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
                        ft.Text(p, size=16, weight=ft.FontWeight.BOLD, overflow=ft.TextOverflow.ELLIPSIS)
                    ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    width=160, height=160,
                    bgcolor=COLOR_SURFACE,
                    border_radius=10,
                    padding=10,
                    on_click=lambda e, name=p: on_project_click(e, name),
                    animate=ft.Animation(200, "easeOut"),
                )
            )

        return ft.View(
            "/",
            [
                ft.Container(
                    content=ft.Column([
                        ft.Text("UNITED GLASS ESTIMATOR", size=30, weight=ft.FontWeight.BOLD, color=COLOR_ACCENT, font_family="Arial"),
                        ft.Text("Select or create a project to begin", size=16, color=COLOR_TEXT_DIM),
                        ft.Divider(color="transparent", height=20),
                        ft.Row([
                            new_proj_name,
                            ft.IconButton(ft.Icons.ADD_CIRCLE, icon_color=COLOR_ACCENT, icon_size=40, on_click=add_project_click, tooltip="Create Project")
                        ]),
                        ft.Divider(color="transparent", height=20),
                        ft.Row(project_tiles, wrap=True, spacing=20, run_spacing=20)
                    ]),
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
            is_yes45 = inputs["system"].value == "YES 45TU FRONT SET(OG)"
            inputs["bays_wide"].parent.visible = is_yes45
            inputs["bays_tall"].parent.visible = is_yes45
            inputs["custom_w"].parent.visible = is_yes45
            inputs["custom_h"].parent.visible = is_yes45
            page.update()

        def on_elevation_load(e):
            elev_name = inputs["saved_elev"].value
            if not elev_name or elev_name not in state["saved_elevations"]: 
                clear_workspace()
                return
            
            data = state["saved_elevations"][elev_name]
            inputs["system"].value = data.get("system", state["system_options"][0])
            inputs["finish"].value = data.get("finish", state["finish_options"][0])
            inputs["type"].value = elev_name
            inputs["count"].value = str(data.get("total_count", ""))
            inputs["width"].value = str(data.get("opening_width_inches", ""))
            inputs["height"].value = str(data.get("opening_height_inches", ""))
            
            inputs["bays_wide"].value = str(data.get("bays_wide", ""))
            inputs["bays_tall"].value = str(data.get("bays_tall", ""))
            inputs["custom_w"].value = ",".join(map(str, data.get("custom_bay_widths", [])))
            inputs["custom_h"].value = ",".join(map(str, data.get("custom_bay_heights", [])))
            
            load_doors(elev_name)
            render_doors()
            update_bay_visibility(None)
            page.update()

        def clear_workspace():
            inputs["type"].value = ""
            inputs["count"].value = ""
            inputs["width"].value = ""
            inputs["height"].value = ""
            inputs["bays_wide"].value = ""
            inputs["bays_tall"].value = ""
            inputs["custom_w"].value = ""
            inputs["custom_h"].value = ""
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
                                ft.Text(f"Door {i+1}", weight="bold", color="white"),
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
            state["current_doors"].pop(idx)
            save_doors_action()
            render_doors()

        def save_doors_action():
            # Saves current doors to file associated with current elevation input name
            elev_name = inputs["type"].value
            if elev_name:
                save_doors(elev_name)

        def modify_door(action):
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
            page.update()

        def save_elevation_action(e):
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
                    
                    # Helper to parse CSV
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
                    data["custom_bay_widths"] = parse_csv(inputs["custom_w"].value, w, bw)
                    data["custom_bay_heights"] = parse_csv(inputs["custom_h"].value, h, bh)
                    
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
                inputs["saved_elev"].options = [ft.dropdown.Option(x) for x in sorted(state["saved_elevations"].keys())]
                inputs["saved_elev"].value = elev
                show_snack("Elevation Saved Successfully", "green")
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
                
                # Update UI
                inputs["saved_elev"].options = [ft.dropdown.Option(x) for x in sorted(state["saved_elevations"].keys())]
                inputs["saved_elev"].value = None
                clear_workspace()
                show_snack("Elevation Deleted", "red")

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
            ft.IconButton(ft.Icons.ARROW_BACK, on_click=lambda e: page.go("/"), icon_color="white"),
            ft.Text(state["current_project"].upper(), size=20, weight="bold", color="white"),
            ft.Container(expand=True),
            ft.ElevatedButton("GENERATE REPORT", bgcolor="white", color="black", on_click=gen_full_report)
        ], height=60, alignment=ft.MainAxisAlignment.START)

        # Left Col: Elevation Form
        elev_options = sorted(state["saved_elevations"].keys())
        saved_dd = create_dropdown("Load Elevation", "saved_elev", elev_options, on_change=on_elevation_load)
        
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
                ft.Row([create_input_field("Bays Wide", "bays_wide"), create_input_field("Bays Tall", "bays_tall")]),
                ft.Row([create_input_field("Custom Widths (csv)", "custom_w"), create_input_field("Custom Heights (csv)", "custom_h")]),
            ]), margin=ft.margin.only(top=10)),

            ft.Container(
                content=ft.Row([
                    ft.ElevatedButton("SAVE ELEVATION", bgcolor=COLOR_ACCENT, color="white", on_click=save_elevation_action, expand=True, height=50),
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
            ft.Row([create_dropdown("Size", "door_size", state["door_options"]), create_input_field("Count", "door_count")]),
            create_dropdown("Style", "door_stile", state["stile_options"]),
            ft.Text("Hardware:", size=12, color=COLOR_TEXT_DIM),
            ft.Column([cb for cb in hardware_cbs.values()], spacing=0),
            ft.Row([
                ft.ElevatedButton("ADD", bgcolor=COLOR_SURFACE, color="white", on_click=lambda e: modify_door("add"), expand=True),
                ft.ElevatedButton("UPDATE", bgcolor=COLOR_SURFACE, color="white", on_click=lambda e: modify_door("update"), expand=True),
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
    ft.app(target=main)
