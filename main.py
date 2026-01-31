import flet as ft
import json, os, re, sys, datetime, base64, io, time, traceback, shutil

# Assuming your utils and systems are in their respective directories
from utils.excel_generator import generate_excel_report
from systems.yes45tu_front_set import calculate_yes45tu_quantities
from utils.formulas import calculate_rectangle_area, calculate_perimeter, calculate_door_info
from utils.waste_calculator import calculate_waste_statistics, get_waste_percentage_color
from utils.database import db  # Centralized database

# PDF export (optional - will fail gracefully if reportlab not installed)
try:
    from utils.pdf_generator import export_project_to_pdf, REPORTLAB_AVAILABLE
except ImportError:
    REPORTLAB_AVAILABLE = False
    print("[WARNING] PDF export not available (reportlab not installed)")

# ML Predictor (optional - will fail gracefully if sklearn not installed)
try:
    from utils.ml_predictor import (
        predict_project_cost, get_pattern_insights, train_ml_model,
        get_training_status, add_project_to_training, collect_training_data_from_projects,
        is_in_training, remove_project_from_training, remove_elevation_from_training
    )
    ML_AVAILABLE = True
except ImportError as e:
    ML_AVAILABLE = False
    print(f"[WARNING] ML predictions not available: {e}")


# --- Constants & Config ---
PROJECTS_DIR = ".files"
MASTER_PROJECT_LIST_FILE = os.path.join(PROJECTS_DIR, "projects_list.json")

# --- PyInstaller Resource Path Helper ---
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Get assets directory path (works for both dev and PyInstaller)
def get_assets_dir():
    """Get the assets directory path"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
        assets_path = os.path.join(base_path, "assets")
        if os.path.exists(assets_path):
            return os.path.abspath(assets_path)  # Return absolute path
    except Exception:
        pass
    # Fallback to absolute path for development
    assets_path = os.path.join(os.path.abspath("."), "assets")
    if os.path.exists(assets_path):
        return os.path.abspath(assets_path)
    # Last resort: return relative path
    return os.path.abspath("assets") if os.path.exists("assets") else "assets"

# United Glass color palette (matching logo)
COLOR_BG = "#000000"       # Black background (matching logo)
COLOR_SURFACE = "#1A1A1A"  # Dark grey card background
COLOR_ACCENT = "#0073E6"   # Vibrant blue (matching logo U)
COLOR_TEXT = "#FFFFFF"     # White text for readability on black
COLOR_TEXT_DIM = "#B3B3B3" # Light grey for secondary text
COLOR_INPUT_BG = "#2A2A2A" # Dark grey input background
COLOR_ACCENT_LIGHT = "#D3D3D3"  # Light grey (matching logo side surfaces)

def create_bay_diagram_base64(bays_wide, bays_tall, opening_width, opening_height, custom_bay_widths=None, custom_bay_heights=None):
    """
    Creates a bay distribution diagram and returns it as a base64 string for Flet display.
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("PIL/Pillow not available, skipping diagram generation")
        return None
    
    if bays_wide <= 0 or bays_tall <= 0 or opening_width <= 0 or opening_height <= 0:
        return None
    
    # Diagram dimensions
    diagram_width = 450
    diagram_height = 350
    margin = 40
    
    # Calculate bay dimensions
    if custom_bay_widths and len(custom_bay_widths) == bays_wide:
        bay_widths = [max(0.1, float(w)) for w in custom_bay_widths]  # Ensure minimum positive value
    else:
        bay_widths = [max(0.1, opening_width / bays_wide)] * bays_wide
    
    if custom_bay_heights and len(custom_bay_heights) == bays_tall:
        bay_heights = [max(0.1, float(h)) for h in custom_bay_heights]  # Ensure minimum positive value
    else:
        bay_heights = [max(0.1, opening_height / bays_tall)] * bays_tall
    
    if not bay_widths or not bay_heights:
        return None
    
    # Validate all dimensions are positive
    if any(w <= 0 for w in bay_widths) or any(h <= 0 for h in bay_heights):
        return None
    
    # Create image with dark background to match app theme
    img = Image.new('RGB', (diagram_width, diagram_height), color='#1A1A1A')
    draw = ImageDraw.Draw(img)
    
    # Calculate scaling to fit in diagram
    max_display_width = diagram_width - 2 * margin
    max_display_height = diagram_height - 2 * margin - 60
    
    total_width = sum(bay_widths)
    total_height = sum(bay_heights)
    scale_x = max_display_width / total_width if total_width > 0 else 1
    scale_y = max_display_height / total_height if total_height > 0 else 1
    scale = min(scale_x, scale_y)
    
    # Calculate starting position (centered)
    scaled_total_width = total_width * scale
    scaled_total_height = total_height * scale
    start_x = margin + (max_display_width - scaled_total_width) / 2
    start_y = margin + 35
    
    # Try to load fonts
    try:
        font_large = ImageFont.truetype("arial.ttf", 14)
        font_small = ImageFont.truetype("arial.ttf", 10)
    except:
        try:
            font_large = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 14)
            font_small = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 10)
        except:
            font_large = ImageFont.load_default()
            font_small = ImageFont.load_default()
    
    # Draw title
    title = f"Bay Distribution ({bays_wide}W x {bays_tall}H)"
    draw.text((diagram_width // 2, 15), title, fill='#0073E6', anchor='mm', font=font_large)
    
    # Add C/L Dimensions note
    cl_note = "C/L Dimensions"
    draw.text((diagram_width // 2, 30), cl_note, fill='#B3B3B3', anchor='mm', font=font_small)
    
    # Draw bays
    current_y = start_y
    bay_num = 1
    for row in range(bays_tall):
        current_x = start_x
        for col in range(bays_wide):
            bay_w = bay_widths[col] * scale
            bay_h = bay_heights[row] * scale
            
            # Ensure minimum size to avoid drawing errors
            bay_w = max(10.0, bay_w)
            bay_h = max(10.0, bay_h)
            
            # Calculate rectangle coordinates
            x0 = current_x
            y0 = current_y
            x1 = current_x + bay_w
            y1 = current_y + bay_h
            
            # Validate coordinates
            if x1 <= x0 or y1 <= y0:
                continue  # Skip invalid rectangles
            
            # Draw bay rectangle with blue border
            draw.rectangle(
                [x0, y0, x1, y1],
                outline='#0073E6',
                width=2
            )
            
            # Draw fill with semi-transparent effect (only if there's enough space)
            if bay_w > 4 and bay_h > 4:
                fill_x0 = x0 + 2
                fill_y0 = y0 + 2
                fill_x1 = x1 - 2
                fill_y1 = y1 - 2
                if fill_x1 > fill_x0 and fill_y1 > fill_y0:
                    draw.rectangle(
                        [fill_x0, fill_y0, fill_x1, fill_y1],
                        fill='#2A2A2A'
                    )
            
            # Bay center
            bay_center_x = current_x + bay_w / 2
            bay_center_y = current_y + bay_h / 2
            
            # Draw bay number
            draw.text((bay_center_x, bay_center_y - 8), f"B{bay_num}", fill='#FFFFFF', anchor='mm', font=font_small)
            
            # Draw dimensions
            dim_text = f"{bay_widths[col]:.1f}\" x {bay_heights[row]:.1f}\""
            draw.text((bay_center_x, bay_center_y + 8), dim_text, fill='#B3B3B3', anchor='mm', font=font_small)
            
            current_x += bay_w
            bay_num += 1
        current_y += bay_h
    
    # Draw overall dimensions
    dim_text = f"Total: {opening_width:.1f}\" W x {opening_height:.1f}\" H"
    draw.text((diagram_width // 2, diagram_height - 25), dim_text, fill='#FFFFFF', anchor='mm', font=font_small)
    
    # Convert to base64
    img_bytes = io.BytesIO()
    img.save(img_bytes, format='PNG')
    img_bytes.seek(0)
    return base64.b64encode(img_bytes.getvalue()).decode('utf-8')

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
        "projects": [], "current_project": None, "saved_elevations": {},
        "current_doors": [], "selected_door_index": None,
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

    # --- Helper Functions ---
    def get_input_pct(key):
        """Get percentage value from input field, returns 0.0 if invalid"""
        field = inputs.get(key)
        if not field: return 0.0
        val = getattr(field, 'value', '') or ''
        try: return float(val) if str(val).strip() else 0.0
        except: return 0.0

    # --- Data Loading Functions (using centralized database) ---
    def load_projects():
        state["projects"] = db.get_projects()

    def save_projects():
        db.save_projects(state["projects"])

    def get_project_paths(project_name):
        """Get paths for Excel reports and temp files."""
        clean_name = project_name.replace(" ", "_").replace("/", "_")
        base = os.path.join(PROJECTS_DIR, clean_name)
        return {
            "excel": f"{base}_Report.xlsx",
            "elevations": f"{base}_Elevations.json",
            "materials": f"{base}_ExtraMaterials.json",
            "settings": f"{base}_Settings.json"
        }
    
    def prepare_temp_files_for_excel(project_name, elevations_override=None):
        """Write database data to temp JSON files for Excel generator. Use elevations_override when provided (e.g. fresh save) to avoid DB read lag."""
        paths = get_project_paths(project_name)
        
        # Write elevations to temp file - prefer in-memory data when provided to ensure latest state
        elevations = elevations_override if elevations_override is not None else db.get_elevations(project_name)
        os.makedirs(os.path.dirname(paths["elevations"]) or ".", exist_ok=True)
        with open(paths["elevations"], 'w') as f:
            json.dump(elevations, f, indent=4)
        
        # Write materials to temp file
        materials = db.get_materials(project_name)
        with open(paths["materials"], 'w') as f:
            json.dump(materials, f, indent=4)
        
        # Write settings to temp file
        settings = db.get_settings(project_name)
        with open(paths["settings"], 'w') as f:
            json.dump(settings, f, indent=4)
        
        return paths
    
    def cleanup_temp_files(paths):
        """Delete temp JSON files after Excel generation."""
        for key in ["elevations", "materials", "settings"]:
            if key in paths and os.path.exists(paths[key]):
                try:
                    os.remove(paths[key])
                except:
                    pass
    
    def sync_from_temp_files(project_name, paths):
        """Read any changes from temp files back to database, then clean up."""
        # The Excel generator may have modified the elevations
        if os.path.exists(paths["elevations"]):
            try:
                with open(paths["elevations"], 'r') as f:
                    elevations = json.load(f)
                db.save_elevations(project_name, elevations)
            except Exception as e:
                print(f"[WARNING] Error syncing elevations: {e}")
        
        # The Excel generator may have modified the extra materials (waste/leftovers)
        if os.path.exists(paths.get("materials")):
            try:
                with open(paths["materials"], 'r') as f:
                    materials = json.load(f)
                db.save_materials(project_name, materials)
                print(f"[OK] Synced {len(materials)} extra materials to database")
            except Exception as e:
                print(f"[WARNING] Error syncing materials: {e}")
        
        # Clean up temp files
        cleanup_temp_files(paths)
    
    def load_project_settings(project_name):
        return db.get_settings(project_name)
    
    def save_project_settings(project_name, settings):
        db.save_settings(project_name, settings)

    def load_elevations(project_name):
        state["saved_elevations"] = db.get_elevations(project_name)

    def load_doors(elev_type):
        if not state["current_project"] or not elev_type:
            state["current_doors"] = []
            return
        state["current_doors"] = db.get_doors(state["current_project"], elev_type)

    def save_doors(elev_type):
        if state["current_project"] and elev_type:
            db.save_doors(state["current_project"], elev_type, state["current_doors"])

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
        
        # ============================================================
        # HELPER: Get project metadata (systems, finishes, elevation count)
        # ============================================================
        def get_project_metadata(project_name):
            """Load project's elevations and extract metadata for filtering."""
            metadata = {
                "systems": set(),
                "finishes": set(),
                "elevation_count": 0
            }
            try:
                elevations = db.get_elevations(project_name)
                metadata["elevation_count"] = len(elevations)
                for elev_name, elev_data in elevations.items():
                    if elev_data.get("system"):
                        metadata["systems"].add(elev_data["system"])
                    if elev_data.get("finish"):
                        metadata["finishes"].add(elev_data["finish"])
            except Exception:
                pass
            return metadata
        
        # ============================================================
        # COLLECT ALL SYSTEMS AND FINISHES FOR FILTER DROPDOWNS
        # ============================================================
        all_systems = set()
        all_finishes = set()
        project_metadata_cache = {}
        
        for p in state["projects"]:
            meta = get_project_metadata(p)
            project_metadata_cache[p] = meta
            all_systems.update(meta["systems"])
            all_finishes.update(meta["finishes"])
        
        # Sort for dropdown display
        all_systems = sorted(list(all_systems))
        all_finishes = sorted(list(all_finishes))
        
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
            """Show confirmation dialog before deleting project"""
            def do_delete(e):
                if name in state["projects"]:
                    # Remove from ML training data if ML is available
                    removed_count = 0
                    if ML_AVAILABLE:
                        try:
                            removed_count = remove_project_from_training(name)
                        except Exception as ex:
                            print(f"[ML] Error removing project from training: {ex}")
                    
                    # Delete project and all its data from database
                    db.delete_project(name)
                    
                    # Update local state
                    state["projects"].remove(name)

                    # Close dialog first
                    dlg.open = False
                    page.update()
                    
                    # Refresh view
                    page.views.pop()
                    page.views.append(build_projects_view())
                    
                    # If on ML Analytics page, refresh it too
                    if len(page.views) > 1 and page.views[-1].route == "/ml_analytics":
                        page.views.pop()
                        page.views.append(build_ml_analytics_view())
                    
                    page.update()
                    
                    # Show snack with ML info if applicable
                    if ML_AVAILABLE and removed_count > 0:
                        show_snack(f"Project '{name}' deleted ({removed_count} training samples removed)", "red")
                    else:
                        show_snack(f"Project '{name}' deleted", "red")
            
            def cancel_delete(e):
                dlg.open = False
                page.update()
            
            dlg = ft.AlertDialog(
                modal=True,
                title=ft.Text("Delete Project", color=COLOR_TEXT),
                content=ft.Text(
                    f"Are you sure you want to delete '{name}'?\n\nThis action cannot be undone.",
                    color=COLOR_TEXT,
                    size=14
                ),
                actions=[
                    ft.TextButton("Cancel", on_click=cancel_delete),
                    ft.TextButton("Delete", on_click=do_delete, style=ft.ButtonStyle(color="#FF5252"))
                ],
                bgcolor=COLOR_SURFACE,
            )
            
            page.overlay.append(dlg)
            dlg.open = True
            page.update()

        # ============================================================
        # SEARCH & FILTER UI COMPONENTS (Professional Design)
        # ============================================================
        search_field = ft.TextField(
            hint_text="Search by project name...",
            prefix_icon=ft.Icons.SEARCH_ROUNDED,
            border_radius=10,
            bgcolor="#0D0D0D",
            border_color="#333333",
            focused_border_color=COLOR_ACCENT,
            color=COLOR_TEXT,
            hint_style=ft.TextStyle(color="#666666", size=13),
            width=300,
            height=48,
            content_padding=ft.padding.only(left=15, right=15, top=12, bottom=12),
            text_size=14,
            cursor_color=COLOR_ACCENT
        )
        
        system_filter = ft.Dropdown(
            hint_text="All Systems",
            options=[ft.dropdown.Option("All Systems")] + [ft.dropdown.Option(s) for s in all_systems],
            value="All Systems",
            width=240,
            bgcolor="#0D0D0D",
            border_color="#333333",
            focused_border_color=COLOR_ACCENT,
            color=COLOR_TEXT,
            hint_style=ft.TextStyle(color="#666666", size=13),
            text_size=13,
            border_radius=10,
            content_padding=ft.padding.symmetric(horizontal=15, vertical=12)
        )
        
        finish_filter = ft.Dropdown(
            hint_text="All Finishes",
            options=[ft.dropdown.Option("All Finishes")] + [ft.dropdown.Option(f) for f in all_finishes],
            value="All Finishes",
            width=180,
            bgcolor="#0D0D0D",
            border_color="#333333",
            focused_border_color=COLOR_ACCENT,
            color=COLOR_TEXT,
            hint_style=ft.TextStyle(color="#666666", size=13),
            text_size=13,
            border_radius=10,
            content_padding=ft.padding.symmetric(horizontal=15, vertical=12)
        )
        
        result_count_text_widget = ft.Text(
            f"Showing {len(state['projects'])} of {len(state['projects'])} projects",
            size=12,
            color=COLOR_TEXT_DIM,
            weight=ft.FontWeight.W_400
        )
        result_count_text = ft.Container(
            content=ft.Row([
                ft.Container(
                    content=ft.Icon(ft.Icons.INVENTORY_2_OUTLINED, size=14, color=COLOR_ACCENT),
                    padding=4
                ),
                result_count_text_widget
            ], spacing=6, alignment=ft.MainAxisAlignment.CENTER),
            padding=ft.padding.symmetric(horizontal=12, vertical=6),
            bgcolor="#0D0D0D",
            border_radius=20
        )
        
        # Container for project tiles (will be updated dynamically)
        projects_grid = ft.Row([], wrap=True, spacing=20, run_spacing=25, alignment=ft.MainAxisAlignment.CENTER)
        
        # ============================================================
        # FILTER LOGIC
        # ============================================================
        def apply_filters(e=None):
            """Apply search and filter criteria to project list."""
            search_term = (search_field.value or "").lower().strip()
            selected_system = system_filter.value
            selected_finish = finish_filter.value
            
            filtered_projects = []
            
            for p in state["projects"]:
                # Search by name
                if search_term and search_term not in p.lower():
                    continue
                
                # Filter by system
                meta = project_metadata_cache.get(p, {"systems": set(), "finishes": set()})
                if selected_system and selected_system != "All Systems":
                    if selected_system not in meta["systems"]:
                        continue
                
                # Filter by finish
                if selected_finish and selected_finish != "All Finishes":
                    if selected_finish not in meta["finishes"]:
                        continue
                
                filtered_projects.append(p)
            
            # Rebuild project tiles with enhanced design
            projects_grid.controls.clear()
            for p in filtered_projects:
                meta = project_metadata_cache.get(p, {"elevation_count": 0, "systems": set(), "finishes": set()})
                elev_count = meta["elevation_count"]
                
                # Create the clickable card using GestureDetector for reliable clicks
                def create_project_card(project_name, elevation_count):
                    return ft.GestureDetector(
                        content=ft.Container(
                            content=ft.Column([
                                # Delete button row (top) - separate click handler
                                ft.Row([
                                    ft.Container(expand=True),
                                    ft.IconButton(
                                        ft.Icons.DELETE_OUTLINE, 
                                        icon_color="#FF5252", 
                                        icon_size=18,
                                        tooltip="Delete Project", 
                                        on_click=lambda e, name=project_name: delete_project_click(e, name),
                                    ),
                                ], alignment=ft.MainAxisAlignment.END),
                                # Icon with gradient background
                                ft.Container(
                                    content=ft.Icon(ft.Icons.FOLDER_OPEN_ROUNDED, size=38, color="#FFFFFF"),
                                    width=65,
                                    height=65,
                                    border_radius=32,
                                    gradient=ft.LinearGradient(
                                        begin=ft.alignment.top_left,
                                        end=ft.alignment.bottom_right,
                                        colors=[COLOR_ACCENT, "#005BB5"]
                                    ),
                                    alignment=ft.alignment.center,
                                ),
                                ft.Container(height=10),
                                # Project name
                                ft.Text(project_name, size=14, weight=ft.FontWeight.W_600, color=COLOR_TEXT, 
                                       overflow=ft.TextOverflow.ELLIPSIS, max_lines=2, 
                                       text_align=ft.TextAlign.CENTER, width=140),
                                # Elevation count with icon
                                ft.Row([
                                    ft.Icon(ft.Icons.LAYERS_OUTLINED, size=14, color=COLOR_TEXT_DIM),
                                    ft.Text(f"{elevation_count} elevation{'s' if elevation_count != 1 else ''}", 
                                           size=11, color=COLOR_TEXT_DIM, weight=ft.FontWeight.W_400),
                                ], spacing=4, alignment=ft.MainAxisAlignment.CENTER),
                            ], alignment=ft.MainAxisAlignment.START, 
                               horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=2),
                            width=175, height=200,
                            bgcolor=COLOR_SURFACE,
                            border_radius=16,
                            padding=ft.padding.only(top=5, bottom=15, left=10, right=10),
                            border=ft.Border(
                                ft.BorderSide(1, "#2A2A2A"),
                                ft.BorderSide(1, "#2A2A2A"),
                                ft.BorderSide(1, "#2A2A2A"),
                                ft.BorderSide(1, "#2A2A2A")
                            ),
                        ),
                        on_tap=lambda e, name=project_name: on_project_click(e, name),
                        mouse_cursor=ft.MouseCursor.CLICK
                    )
                
                projects_grid.controls.append(create_project_card(p, elev_count))
            
            # Update result count
            result_count_text_widget.value = f"Showing {len(filtered_projects)} of {len(state['projects'])} projects"
            
            page.update()
        
        # Attach filter handlers
        search_field.on_change = apply_filters
        system_filter.on_change = apply_filters
        finish_filter.on_change = apply_filters
        
        def clear_filters(e):
            """Reset all filters."""
            search_field.value = ""
            system_filter.value = "All Systems"
            finish_filter.value = "All Finishes"
            apply_filters()
        
        # Initial load of projects
        apply_filters()
        
        # Custom styled input for new project
        new_proj_name = ft.TextField(
            hint_text="Enter project name...",
            border_radius=8,
            bgcolor="#0D0D0D",
            border_color="#333333",
            focused_border_color=COLOR_ACCENT,
            color=COLOR_TEXT,
            hint_style=ft.TextStyle(color="#666666", size=13),
            width=280,
            height=42,
            content_padding=ft.padding.symmetric(horizontal=15, vertical=10),
            text_size=14,
            cursor_color=COLOR_ACCENT
        )

        # Try to load the United Glass logo
        # Try multiple possible filenames
        logo_filenames = ["R.png", "united_glass_logo.png"]
        logo_image = None
        
        for logo_filename in logo_filenames:
            # Try resource_path first (for PyInstaller bundle)
            try:
                logo_path = resource_path(os.path.join("assets", logo_filename))
                if os.path.exists(logo_path):
                    logo_image = ft.Image(
                        src=logo_path,
                        width=200,
                        height=200,
                        fit=ft.ImageFit.CONTAIN,
                    )
                    break  # Successfully loaded, exit loop
            except Exception:
                pass
            
            # Try relative path (for development)
            try:
                logo_path = os.path.join("assets", logo_filename)
                if os.path.exists(logo_path):
                    # Try with just filename (since assets_dir is set in Flet)
                    logo_image = ft.Image(
                        src=logo_filename,
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
        
        # Empty state for when no projects exist
        empty_state = ft.Container(
            content=ft.Column([
                ft.Icon(ft.Icons.FOLDER_OFF_OUTLINED, size=80, color=COLOR_TEXT_DIM),
                ft.Container(height=16),
                ft.Text("No Projects Yet", size=24, weight=ft.FontWeight.W_600, color=COLOR_TEXT),
                ft.Text("Create your first project to get started", size=14, color=COLOR_TEXT_DIM),
            ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            padding=60,
            visible=len(state["projects"]) == 0
        )
        
        return ft.View(
            "/",
            [
                ft.Container(
                    content=ft.Column([
                        # ========== HEADER SECTION ==========
                        ft.Container(
                            content=ft.Row([
                                # Logo
                                ft.Container(
                                    content=logo_display,
                                    alignment=ft.alignment.center,
                                ),
                                ft.Container(width=30),
                                # Title and subtitle
                                ft.Column([
                                    ft.Text("ESTIMATION TOOL", 
                                           size=42, 
                                           weight=ft.FontWeight.BOLD, 
                                           color=COLOR_ACCENT,
                                           font_family="Arial"),
                                    ft.Text("Professional Storefront & Curtain Wall Estimator", 
                                           size=16, 
                                           color=COLOR_TEXT_DIM,
                                           weight=ft.FontWeight.W_400),
                                ], spacing=4, alignment=ft.MainAxisAlignment.CENTER),
                            ], alignment=ft.MainAxisAlignment.CENTER),
                            padding=ft.padding.only(bottom=30),
                        ),
                        
                        # ========== ACTION BAR (Create + ML Analytics) ==========
                        ft.Container(
                            content=ft.Row([
                                # Create Project Card
                                ft.Container(
                                    content=ft.Row([
                                        ft.Icon(ft.Icons.CREATE_NEW_FOLDER_OUTLINED, size=24, color=COLOR_ACCENT),
                                        ft.Container(width=12),
                                        new_proj_name,
                                        ft.Container(width=8),
                                        ft.ElevatedButton(
                                            "Create",
                                            icon=ft.Icons.ADD,
                                            bgcolor=COLOR_ACCENT,
                                            color="white",
                                            on_click=add_project_click,
                                            height=42,
                                            style=ft.ButtonStyle(
                                                shape=ft.RoundedRectangleBorder(radius=8),
                                                elevation=0
                                            )
                                        ),
                                    ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                                    bgcolor=COLOR_SURFACE,
                                    padding=ft.padding.symmetric(horizontal=20, vertical=12),
                                    border_radius=12,
                                    border=ft.Border(
                                        ft.BorderSide(1, "#2A2A2A"),
                                        ft.BorderSide(1, "#2A2A2A"),
                                        ft.BorderSide(1, "#2A2A2A"),
                                        ft.BorderSide(1, "#2A2A2A")
                                    ),
                                    expand=True
                                ),
                                ft.Container(width=15),
                                # ML Analytics Button
                                ft.Container(
                                    content=ft.ElevatedButton(
                                        "ML Analytics",
                                        icon=ft.Icons.AUTO_GRAPH,
                                        bgcolor="#4CAF50" if ML_AVAILABLE else "#2A2A2A",
                                        color="white" if ML_AVAILABLE else COLOR_TEXT_DIM,
                                        on_click=lambda e: page.go("/ml_analytics") if ML_AVAILABLE else show_snack("ML not available. Install scikit-learn.", "red"),
                                        height=42,
                                        style=ft.ButtonStyle(
                                            shape=ft.RoundedRectangleBorder(radius=8),
                                            elevation=0
                                        ),
                                        tooltip="Machine Learning Analytics - Predict costs from historical data"
                                    ),
                                ),
                            ], alignment=ft.MainAxisAlignment.CENTER),
                            padding=ft.padding.only(bottom=25),
                        ),
                        
                        # ========== SEARCH & FILTER SECTION ==========
                        ft.Container(
                            content=ft.Column([
                                # Section Header
                                ft.Row([
                                    ft.Row([
                                        ft.Icon(ft.Icons.FILTER_LIST_ROUNDED, size=22, color=COLOR_ACCENT),
                                        ft.Container(width=8),
                                        ft.Text("Search & Filter", size=16, weight=ft.FontWeight.W_600, color=COLOR_TEXT),
                                    ]),
                                    ft.Container(expand=True),
                                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                                ft.Container(height=15),
                                # Filter Controls
                                ft.Row([
                                    search_field,
                                    system_filter,
                                    finish_filter,
                                ], spacing=20, alignment=ft.MainAxisAlignment.CENTER, wrap=True),
                                ft.Container(height=12),
                                # Result count
                                result_count_text,
                            ], horizontal_alignment=ft.CrossAxisAlignment.STRETCH, spacing=0),
                            bgcolor=COLOR_SURFACE,
                            padding=20,
                            border_radius=14,
                            border=ft.Border(
                                ft.BorderSide(1, "#2A2A2A"),
                                ft.BorderSide(1, "#2A2A2A"),
                                ft.BorderSide(1, "#2A2A2A"),
                                ft.BorderSide(1, "#2A2A2A")
                            ),
                        ),
                        ft.Container(height=25),
                        
                        # ========== PROJECTS SECTION HEADER ==========
                        ft.Row([
                            ft.Row([
                                ft.Icon(ft.Icons.FOLDER_COPY_OUTLINED, size=22, color=COLOR_ACCENT),
                                ft.Container(width=8),
                                ft.Text("Your Projects", size=18, weight=ft.FontWeight.W_600, color=COLOR_TEXT),
                            ]),
                            ft.Container(expand=True),
                            ft.Text(f"{len(state['projects'])} total", size=13, color=COLOR_TEXT_DIM),
                        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                        ft.Container(height=15),
                        
                        # ========== PROJECTS GRID ==========
                        empty_state,
                        projects_grid
                        
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, scroll=ft.ScrollMode.AUTO),
                    padding=ft.padding.symmetric(horizontal=50, vertical=30),
                    alignment=ft.alignment.top_center,
                    expand=True
                )
            ],
            bgcolor=COLOR_BG,
            padding=0
        )

    def build_ml_analytics_view():
        """Build the ML Analytics view - accessible from projects page."""
        if not ML_AVAILABLE:
            return ft.View(
                "/ml_analytics",
                [ft.Container(
                    content=ft.Column([
                        ft.Text("ML Analytics Not Available", size=24, weight="bold", color=COLOR_ACCENT),
                        ft.Text("Please install scikit-learn: pip install scikit-learn numpy", color=COLOR_TEXT_DIM),
                        ft.ElevatedButton("Back to Projects", on_click=lambda e: page.go("/"))
                    ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    expand=True,
                    alignment=ft.alignment.center
                )],
                bgcolor=COLOR_BG
            )
        
        # ML UI components - start with empty state
        ml_status_text = ft.Text("Ready to load projects", size=14, color=COLOR_TEXT_DIM)
        ml_minimum_warning = ft.Container(
            visible=False,
            content=ft.Row([
                ft.Icon(ft.Icons.WARNING, color="#FFC107", size=20),
                ft.Text("", size=13, color="#FFC107", expand=True, weight="bold")
            ], spacing=10),
            bgcolor="#FFF3CD",
            padding=12,
            border_radius=8,
            border=ft.Border(
                ft.BorderSide(1, "#FFC107"),
                ft.BorderSide(1, "#FFC107"),
                ft.BorderSide(1, "#FFC107"),
                ft.BorderSide(1, "#FFC107")
            )
        )
        ml_stats_column = ft.Column([], spacing=10)
        ml_projects_list = ft.Column([], spacing=8, scroll=ft.ScrollMode.AUTO)
        ml_insights_list = ft.Column([], spacing=8)
        
        # Track selected elevations (dict: project_name-elev_name -> elevation_data)
        selected_elevations = {}
        
        # Container for projects list with animation
        ml_projects_container = ft.Container(
            content=ml_projects_list,
            height=500,
            border=ft.Border(
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT)
            ),
            border_radius=8,
            padding=10,
            opacity=0,
            animate_opacity=300
        )
        
        # Container for insights with animation
        ml_insights_container = ft.Container(
            content=ml_insights_list,
            height=500,
            border=ft.Border(
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT)
            ),
            border_radius=8,
            padding=10,
            opacity=0,
            animate_opacity=300
        )
        
        # Empty state for projects
        projects_empty_state = ft.Container(
            content=ft.Column([
                ft.Icon(ft.Icons.AUTO_GRAPH, size=64, color=COLOR_TEXT_DIM),
                ft.Text("No projects loaded", size=18, weight="bold", color=COLOR_TEXT_DIM),
                ft.Text("Click 'Load Projects' to see predictions", size=12, color=COLOR_TEXT_DIM),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=10),
            alignment=ft.alignment.center,
            expand=True
        )
        
        # Empty state for insights
        insights_empty_state = ft.Container(
            content=ft.Column([
                ft.Icon(ft.Icons.INSIGHTS, size=64, color=COLOR_TEXT_DIM),
                ft.Text("No insights available", size=18, weight="bold", color=COLOR_TEXT_DIM),
                ft.Text("Train the model to see pattern insights", size=12, color=COLOR_TEXT_DIM),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=10),
            alignment=ft.alignment.center,
            expand=True
        )
        
        # Track loaded state
        projects_loaded = False
        insights_loaded = False
        
        def refresh_ml_status():
            """Refresh ML training status."""
            status = get_training_status()
            sample_count = status['training_samples']
            ml_status_text.value = f"Model Trained: {'Yes' if status['is_trained'] else 'No'} | Training Samples: {sample_count} | sklearn: {'OK' if status['sklearn_available'] else 'Missing'}"
            ml_status_text.color = "#4CAF50" if status['is_trained'] else COLOR_TEXT_DIM
            
            # Show minimum warning if less than 3 samples
            if sample_count < 3:
                ml_minimum_warning.content.controls[1].value = f"Need minimum of 3 samples to train. Currently have {sample_count} sample(s)."
                ml_minimum_warning.visible = True
            else:
                ml_minimum_warning.visible = False
            
            page.update()
        
        def train_model_action(e):
            """Train ML model from all projects."""
            show_snack("Training model...", COLOR_ACCENT)
            page.update()
            
            success, message = train_ml_model()
            
            if success:
                show_snack(message, "green")
                refresh_ml_status()
                refresh_all_projects()
                show_patterns()
            else:
                show_snack(message, "red")
                refresh_ml_status()
        
        def add_single_to_training(elev_data, project_name, elev_name, button_ref, container_ref=None):
            """Add a single elevation to training data with visual feedback."""
            
            # Show loading state
            if button_ref:
                button_ref.icon = ft.Icons.HOURGLASS_EMPTY
                button_ref.icon_color = COLOR_TEXT_DIM
                button_ref.disabled = True
            page.update()
            
            try:
                result = add_project_to_training(elev_data, project_name, elev_name)
                
                if result:
                    # Show success state - keep selected state
                    if button_ref:
                        button_ref.icon = ft.Icons.CHECK_CIRCLE
                        button_ref.icon_color = "#4CAF50"
                        button_ref.disabled = False
                        button_ref.tooltip = "In training data"
                    
                    # Update container background to show selected
                    if container_ref:
                        container_ref.bgcolor = "#E8F5E9"  # Light green background
                        container_ref.border = ft.Border(
                            ft.BorderSide(2, "#4CAF50"),
                            ft.BorderSide(2, "#4CAF50"),
                            ft.BorderSide(2, "#4CAF50"),
                            ft.BorderSide(2, "#4CAF50")
                        )
                    
                    show_snack("Added to training data!", "green")
                    refresh_ml_status()
                    page.update()
                    
                    # Refresh to show all selected states
                    refresh_all_projects()
                else:
                    # Show already added state
                    if button_ref:
                        button_ref.icon = ft.Icons.CHECK_CIRCLE
                        button_ref.icon_color = "#4CAF50"
                        button_ref.disabled = False
                        button_ref.tooltip = "Already in training data"
                    
                    # Update container to show selected
                    if container_ref:
                        container_ref.bgcolor = "#E8F5E9"
                        container_ref.border = ft.Border(
                            ft.BorderSide(2, "#4CAF50"),
                            ft.BorderSide(2, "#4CAF50"),
                            ft.BorderSide(2, "#4CAF50"),
                            ft.BorderSide(2, "#4CAF50")
                        )
                    
                    show_snack("Already in training data", COLOR_TEXT_DIM)
                    page.update()
                
            except Exception as ex:
                print(f"[ERROR] Error adding to training: {ex}")
                if button_ref:
                    button_ref.icon = ft.Icons.ADD_CIRCLE_OUTLINE
                    button_ref.icon_color = COLOR_ACCENT
                    button_ref.disabled = False
                show_snack(f"Error: {ex}", "red")
                page.update()
        
        def refresh_all_projects():
            """Show all projects with their ML predictions and checkboxes."""
            nonlocal projects_loaded
            
            ml_projects_list.controls = []
            selected_elevations.clear()
            
            # Get all projects
            for project_name in state["projects"]:
                try:
                    elevations = db.get_elevations(project_name)
                    
                    for elev_name, elev_data in elevations.items():
                        prediction = predict_project_cost(elev_data)
                        
                        # Check if already in training
                        in_training = is_in_training(elev_data, project_name, elev_name)
                        
                        # Create unique key for this elevation
                        elev_key = f"{project_name}-{elev_name}"
                        
                        # Create checkbox with data stored
                        elevation_checkbox = ft.Checkbox(
                            value=in_training,  # Pre-check if in training
                            on_change=lambda e, key=elev_key, data=elev_data, proj=project_name, en=elev_name: on_checkbox_change(e, key, data, proj, en),
                            tooltip="Select to add/remove from training data",
                            data={'key': elev_key, 'data': elev_data, 'project': project_name, 'elev_name': elev_name}  # Store data in checkbox
                        )
                        
                        # Color based on confidence
                        conf_color = "#4CAF50" if prediction['confidence'] >= 0.7 else "#FFC107" if prediction['confidence'] >= 0.5 else "#F44336"
                        
                        # Create button with reference for visual feedback (legacy, but keep for quick add)
                        add_button = ft.IconButton(
                            icon=ft.Icons.CHECK_CIRCLE if in_training else ft.Icons.ADD_CIRCLE_OUTLINE,
                            icon_color="#4CAF50" if in_training else COLOR_ACCENT,
                            tooltip="In training data" if in_training else "Add to training data",
                            style=ft.ButtonStyle(
                                shape=ft.CircleBorder(),
                            )
                        )
                        
                        # Create container with reference for styling
                        project_container = ft.Container(
                            content=ft.Row([
                                elevation_checkbox,
                                ft.Column([
                                    ft.Row([
                                        ft.Text(f"{project_name} - {elev_name}", size=13, weight="bold", color=COLOR_TEXT, expand=True),
                                        ft.Container(
                                            content=ft.Icon(ft.Icons.CHECK_CIRCLE, size=16, color="#4CAF50"),
                                            visible=in_training,
                                            tooltip="In training data"
                                        )
                                    ], spacing=5),
                                    ft.Text(
                                        f"{elev_data.get('opening_width_inches', 0):.0f}\" x {elev_data.get('opening_height_inches', 0):.0f}\" | "
                                        f"{elev_data.get('bays_wide', 0)}x{elev_data.get('bays_tall', 0)} bays | "
                                        f"{elev_data.get('finish', 'N/A')}",
                                        size=11, color=COLOR_TEXT_DIM
                                    )
                                ], expand=True, spacing=2),
                                ft.Column([
                                    ft.Text(f"${prediction['predicted_cost']:,.2f}", size=14, weight="bold", color=conf_color),
                                    ft.Text(f"{prediction['confidence']*100:.0f}% confidence", size=10, color=COLOR_TEXT_DIM)
                                ], horizontal_alignment=ft.CrossAxisAlignment.END, spacing=2),
                                add_button
                            ], spacing=10),
                            padding=12,
                            bgcolor="#E8F5E9" if in_training else COLOR_INPUT_BG,
                            border=ft.Border(
                                ft.BorderSide(2, "#4CAF50"),
                                ft.BorderSide(2, "#4CAF50"),
                                ft.BorderSide(2, "#4CAF50"),
                                ft.BorderSide(2, "#4CAF50")
                            ) if in_training else ft.Border(
                                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                                ft.BorderSide(1, COLOR_ACCENT_LIGHT)
                            ),
                            border_radius=8,
                            animate=ft.Animation(300, "easeOut")
                        )
                        
                        # Set up click handler with references
                        add_button.on_click = lambda e, ed=elev_data, pn=project_name, en=elev_name, btn=add_button, cont=project_container: add_single_to_training(ed, pn, en, btn, cont)
                        
                        # Track if in training in selected_elevations
                        if in_training:
                            selected_elevations[elev_key] = {'data': elev_data, 'project': project_name, 'elev_name': elev_name}
                        
                        ml_projects_list.controls.append(project_container)
                except Exception as ex:
                        print(f"[WARNING] Error loading project {project_name}: {ex}")
            
            # Show empty state or content
            if not ml_projects_list.controls:
                ml_projects_container.content = projects_empty_state
                ml_projects_container.opacity = 1
                projects_loaded = False
            else:
                ml_projects_container.content = ml_projects_list
                ml_projects_container.opacity = 1
                projects_loaded = True
            
            page.update()
        
        def on_checkbox_change(e, elev_key, elev_data, project_name, elev_name):
            """Handle checkbox state change."""
            if e.control.value:
                selected_elevations[elev_key] = {'data': elev_data, 'project': project_name, 'elev_name': elev_name}
            else:
                if elev_key in selected_elevations:
                    del selected_elevations[elev_key]
            page.update()
        
        def select_all_projects(e):
            """Select all elevations for training."""
            selected_elevations.clear()
            for container in ml_projects_list.controls:
                if isinstance(container.content, ft.Row) and len(container.content.controls) >= 2:
                    checkbox = container.content.controls[0]
                    if isinstance(checkbox, ft.Checkbox) and checkbox.data:
                        checkbox.value = True
                        # Add to selected using stored data
                        cb_data = checkbox.data
                        selected_elevations[cb_data['key']] = {
                            'data': cb_data['data'],
                            'project': cb_data['project'],
                            'elev_name': cb_data.get('elev_name', '')
                        }
            page.update()
            show_snack("All elevations selected", "green")
        
        def deselect_all_projects(e):
            """Deselect all projects."""
            selected_elevations.clear()
            for container in ml_projects_list.controls:
                if isinstance(container.content, ft.Row):
                    checkbox = container.content.controls[0]
                    if isinstance(checkbox, ft.Checkbox):
                        checkbox.value = False
            page.update()
            show_snack("All elevations deselected", COLOR_TEXT_DIM)
        
        def add_selected_to_training(e):
            """Add all selected elevations to training data."""
            added = 0
            already_in = 0
            for elev_key, info in selected_elevations.items():
                result = add_project_to_training(info['data'], info['project'], info.get('elev_name', ''))
                if result:
                    added += 1
                else:
                    already_in += 1
            
            if added > 0:
                show_snack(f"Added {added} elevation(s) to training data!", "green")
                refresh_ml_status()
                refresh_all_projects()
            elif already_in > 0:
                show_snack(f"All selected elevations already in training data", COLOR_TEXT_DIM)
            else:
                show_snack("No elevations selected", "red")
        
        def remove_selected_from_training(e):
            """Remove all selected elevations from training data."""
            if not selected_elevations:
                show_snack("No elevations selected", "red")
                return
            
            removed = 0
            not_in_training = 0
            for elev_key, info in selected_elevations.items():
                if remove_elevation_from_training(info['data'], info['project'], info.get('elev_name', '')):
                    removed += 1
                else:
                    not_in_training += 1
            
            if removed > 0:
                show_snack(f"Removed {removed} elevation(s) from training data!", "green")
                refresh_ml_status()
                refresh_all_projects()
            elif not_in_training > 0:
                show_snack(f"Selected elevations not in training data", COLOR_TEXT_DIM)
            else:
                show_snack("No elevations selected", "red")
        
        def add_all_to_training(e):
            """Add all projects to training data."""
            added = 0
            for project_name in state["projects"]:
                try:
                    elevations = db.get_elevations(project_name)
                    
                    for elev_name, elev_data in elevations.items():
                        if add_project_to_training(elev_data, project_name, elev_name):
                            added += 1
                except:
                    pass
            
            if added > 0:
                show_snack(f"Added {added} elevations to training data!", "green")
                refresh_ml_status()
            else:
                show_snack("No new data to add", COLOR_TEXT_DIM)
        
        def show_patterns():
            """Show pattern analysis - only shows content when model is trained."""
            nonlocal insights_loaded
            
            ml_insights_list.controls = []
            
            patterns = get_pattern_insights()
            
            # Only show patterns if model is trained
            if patterns.get('error') or not patterns.get('is_trained'):
                ml_insights_container.content = insights_empty_state
                ml_insights_container.opacity = 1
                insights_loaded = False
                page.update()
                return
            
            # Model is trained - show statistics
            if patterns.get('sample_count', 0) > 0:
                ml_insights_list.controls.append(
                    ft.Container(
                        content=ft.Column([
                            ft.Text("Statistics", size=14, weight="bold", color=COLOR_ACCENT),
                            ft.Row([
                                ft.Column([
                                    ft.Text("Avg Cost", size=11, color=COLOR_TEXT_DIM),
                                    ft.Text(f"${patterns.get('avg_cost', 0):,.2f}", size=16, weight="bold", color=COLOR_TEXT)
                                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, expand=True),
                                ft.Column([
                                    ft.Text("Avg Sqft", size=11, color=COLOR_TEXT_DIM),
                                    ft.Text(f"{patterns.get('avg_sqft', 0):.0f} sqft", size=16, weight="bold", color=COLOR_TEXT)
                                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, expand=True),
                                ft.Column([
                                    ft.Text("Avg Size", size=11, color=COLOR_TEXT_DIM),
                                    ft.Text(f"{patterns.get('avg_width', 0):.0f}\" x {patterns.get('avg_height', 0):.0f}\"", size=14, weight="bold", color=COLOR_TEXT)
                                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, expand=True),
                            ], alignment=ft.MainAxisAlignment.SPACE_AROUND)
                        ], spacing=10),
                        padding=15,
                        bgcolor=COLOR_INPUT_BG,
                        border_radius=8
                    )
                )
                
                # Cost range
                ml_insights_list.controls.append(
                    ft.Container(
                        content=ft.Column([
                            ft.Text("Cost Range", size=14, weight="bold", color=COLOR_ACCENT),
                            ft.Row([
                                ft.Text(f"Min: ${patterns.get('min_cost', 0):,.2f}", size=12, color=COLOR_TEXT),
                                ft.Text(" - ", color=COLOR_TEXT_DIM),
                                ft.Text(f"Max: ${patterns.get('max_cost', 0):,.2f}", size=12, color=COLOR_TEXT),
                            ], alignment=ft.MainAxisAlignment.CENTER)
                        ], spacing=10, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                        padding=15,
                        bgcolor=COLOR_INPUT_BG,
                        border_radius=8
                    )
                )
            
            # Common configurations
            configs = patterns.get('common_configurations', [])
            if configs:
                config_items = []
                for c in configs[:5]:
                    config_items.append(
                        ft.Row([
                            ft.Text(f"{c['config']} bays", size=12, color=COLOR_TEXT, expand=True),
                            ft.Text(f"{c['count']} projects ({c['percentage']:.0f}%)", size=11, color=COLOR_TEXT_DIM)
                        ])
                    )
                
                ml_insights_list.controls.append(
                    ft.Container(
                        content=ft.Column([
                            ft.Text("Common Configurations", size=14, weight="bold", color=COLOR_ACCENT),
                            ft.Column(config_items, spacing=5)
                        ], spacing=10),
                        padding=15,
                        bgcolor=COLOR_INPUT_BG,
                        border_radius=8
                    )
                )
            
            # Model status (always show when trained)
            ml_insights_list.controls.append(
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.CHECK_CIRCLE, color="#4CAF50", size=18),
                        ft.Text(
                            f"Model trained - {patterns.get('sample_count', 0)} samples", 
                            size=12, color=COLOR_TEXT
                        )
                    ], spacing=10),
                    padding=10,
                    bgcolor=COLOR_INPUT_BG,
                    border_radius=8
                )
            )
            
            # Show insights
            ml_insights_container.content = ml_insights_list
            ml_insights_container.opacity = 1
            insights_loaded = True
            
            page.update()
        
        def load_projects_action(e):
            """Load projects and show them with animation."""
            show_snack("Loading projects...", COLOR_ACCENT)
            refresh_ml_status()
            refresh_all_projects()
            show_patterns()
            show_snack("Projects loaded!", "green")
        
        # Initialize - start blank
        refresh_ml_status()
        # Don't auto-load projects - user must click "Load Projects"
        ml_projects_container.content = projects_empty_state
        ml_projects_container.opacity = 1
        ml_insights_container.content = insights_empty_state
        ml_insights_container.opacity = 1
        projects_loaded = False
        insights_loaded = False
        
        return ft.View(
            "/ml_analytics",
            [
                ft.Container(
                    content=ft.Column([
                        # Header
                        ft.Row([
                            ft.IconButton(ft.Icons.ARROW_BACK, icon_color=COLOR_ACCENT, on_click=lambda e: page.go("/")),
                            ft.Text("ML ANALYTICS", size=28, weight="bold", color=COLOR_ACCENT, expand=True),
                            ft.ElevatedButton(
                                "Load Projects",
                                icon=ft.Icons.REFRESH,
                                bgcolor=COLOR_ACCENT,
                                color="white",
                                on_click=load_projects_action,
                                height=40
                            ),
                            ft.ElevatedButton(
                                "Train Model",
                                icon=ft.Icons.MODEL_TRAINING,
                                bgcolor="#4CAF50",
                                color="white",
                                on_click=train_model_action,
                                height=40
                            )
                        ], alignment=ft.MainAxisAlignment.START),
                        ml_status_text,
                        ml_minimum_warning,  # Show minimum requirement warning
                        ft.Divider(height=20, color=COLOR_SURFACE),
                        
                        # Main content in two columns
                        ft.Row([
                            # Left: Projects list
                            ft.Container(
                                content=ft.Column([
                                    ft.Row([
                                        ft.Text("PROJECT PREDICTIONS", size=16, weight="bold", color=COLOR_ACCENT, expand=True),
                                    ]),
                                    # Bulk action buttons
                                    ft.Row([
                                        ft.ElevatedButton(
                                            "Select All",
                                            icon=ft.Icons.CHECKLIST,
                                            bgcolor="#2196F3",
                                            color="white",
                                            on_click=select_all_projects,
                                            height=35,
                                            tooltip="Select all elevations"
                                        ),
                                        ft.ElevatedButton(
                                            "Deselect All",
                                            icon=ft.Icons.CLEAR_ALL,
                                            bgcolor=COLOR_TEXT_DIM,
                                            color="white",
                                            on_click=deselect_all_projects,
                                            height=35,
                                            tooltip="Deselect all elevations"
                                        ),
                                        ft.ElevatedButton(
                                            "Add Selected",
                                            icon=ft.Icons.ADD_CIRCLE,
                                            bgcolor="#4CAF50",
                                            color="white",
                                            on_click=add_selected_to_training,
                                            height=35,
                                            tooltip="Add selected elevations to training"
                                        ),
                                        ft.ElevatedButton(
                                            "Remove Selected",
                                            icon=ft.Icons.REMOVE_CIRCLE,
                                            bgcolor="#F44336",
                                            color="white",
                                            on_click=remove_selected_from_training,
                                            height=35,
                                            tooltip="Remove selected elevations from training"
                                        )
                                    ], spacing=8),
                                    ft.Container(height=10),
                                    ml_projects_container
                                ]),
                                expand=True,
                                padding=20,
                                bgcolor=COLOR_SURFACE,
                                border_radius=12,
                                animate=ft.Animation(300, "easeOut")
                            ),
                            
                            # Right: Insights
                            ft.Container(
                                content=ft.Column([
                                    ft.Row([
                                        ft.Text("PATTERN INSIGHTS", size=16, weight="bold", color=COLOR_ACCENT, expand=True),
                                        ft.Container(
                                            content=ft.Row([
                                                ft.Icon(ft.Icons.INSIGHTS, size=16, color=COLOR_TEXT_DIM),
                                                ft.Text("Train model to see insights", size=10, color=COLOR_TEXT_DIM)
                                            ], spacing=5),
                                            tooltip="Patterns learned from training data"
                                        )
                                    ]),
                                    ft.Container(height=10),
                                    ml_insights_container
                                ]),
                                expand=True,
                                padding=20,
                                bgcolor=COLOR_SURFACE,
                                border_radius=12,
                                animate=ft.Animation(300, "easeOut")
                            )
                        ], spacing=20, expand=True)
                    ], expand=True),
                    padding=30,
                    expand=True
                )
            ],
            bgcolor=COLOR_BG,
            padding=0,
            scroll=ft.ScrollMode.AUTO
        )

    def build_workspace_view():
        # --- Event Handlers ---
        def update_door_only_visibility(e):
            """When door_only checked: hide system, finish, quantity, dimensions, bays. Still requires elevation name."""
            door_only = inputs.get("door_only_checkbox") and inputs["door_only_checkbox"].value
            if inputs.get("door_only_rest_container"):
                inputs["door_only_rest_container"].visible = not door_only
            if inputs.get("bay_diagram_container"):
                inputs["bay_diagram_container"].visible = not door_only
            if not door_only:
                update_bay_visibility(e)
            if page.views:
                page.update()

        def update_bay_visibility(e):
            # Guard against unmounted inputs
            if not inputs.get("bays_wide") or not inputs["bays_wide"].parent:
                return

            door_only = inputs.get("door_only_checkbox") and inputs["door_only_checkbox"].value
            is_yes45 = inputs["system"].value == "YES 45TU FRONT SET(OG)"
            show_bays = is_yes45 and not door_only
            if inputs.get("bay_config_container"):
                inputs["bay_config_container"].visible = show_bays
            inputs["bays_wide"].parent.visible = show_bays
            inputs["bays_tall"].parent.visible = show_bays
            if inputs.get("custom_w_container"): inputs["custom_w_container"].visible = show_bays
            if inputs.get("custom_h_container"): inputs["custom_h_container"].visible = show_bays
            
            if is_yes45 and not door_only:
                update_dynamic_bay_inputs(None)
                # Auto-refresh bay diagram after updating inputs
                auto_refresh_bay_diagram()
            else:
                # Hide bay diagram if not YES 45TU or door-only
                if inputs.get("bay_diagram_container"):
                    inputs["bay_diagram_container"].visible = False
            
            if page.views: 
                page.update()

        def update_qty_per_elev_visibility(e):
            # Show/hide the "Quantity Per Elevation" and cost per elevation checkboxes based on count value
            if not inputs.get("qty_per_elev_container"):
                return
            
            try:
                count_val = int(inputs["count"].value) if inputs["count"].value else 0
                should_show = count_val > 1
                inputs["qty_per_elev_container"].visible = should_show
                if inputs.get("cost_per_elev_container"):
                    inputs["cost_per_elev_container"].visible = should_show
                if page.views:
                    page.update()
            except (ValueError, TypeError):
                inputs["qty_per_elev_container"].visible = False
                if inputs.get("cost_per_elev_container"):
                    inputs["cost_per_elev_container"].visible = False
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
                    field = create_input_field(f"Bay {i+1} Width", f"bay_w_{i}", expand=True, value=val, on_change=lambda e: auto_refresh_bay_diagram())
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
                    
                    # Refresh bay diagram after auto-fill
                    auto_refresh_bay_diagram()

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
                    field = create_input_field(f"Bay {i+1} Height", f"bay_h_{i}", expand=True, value=val, on_change=lambda e: auto_refresh_bay_diagram())
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
                    
                    # Refresh bay diagram after auto-fill
                    auto_refresh_bay_diagram()

                inputs["custom_h_col"].controls.append(
                     ft.ElevatedButton("Auto-Fill Remaining Heights", on_click=auto_fill_h, bgcolor=COLOR_ACCENT, color="white")
                )

            inputs["dynamic_h_fields"] = new_h_fields
            
            # Auto-refresh bay diagram when inputs change
            auto_refresh_bay_diagram()
            
            if page.views:
                page.update()
        
        def auto_refresh_bay_diagram():
            """Automatically refresh bay diagram when inputs change."""
            try:
                # Only refresh if YES 45TU system is selected
                if inputs["system"].value != "YES 45TU FRONT SET(OG)":
                    if inputs.get("bay_diagram_container"):
                        inputs["bay_diagram_container"].visible = False
                    return
                
                try:
                    bw = int(inputs["bays_wide"].value) if inputs["bays_wide"].value else 0
                    bh = int(inputs["bays_tall"].value) if inputs["bays_tall"].value else 0
                except:
                    bw, bh = 0, 0
                
                if bw <= 0 or bh <= 0:
                    if inputs.get("bay_diagram_container"):
                        inputs["bay_diagram_container"].visible = False
                    return
                
                # Gather custom widths - collect all filled values
                custom_w = []
                for f in inputs.get("dynamic_w_fields", []):
                    try:
                        if f.value and str(f.value).strip():
                            custom_w.append(float(f.value))
                        else:
                            custom_w.append(None)  # Mark as empty
                    except:
                        custom_w.append(None)
                
                # Gather custom heights - collect all filled values
                custom_h = []
                for f in inputs.get("dynamic_h_fields", []):
                    try:
                        if f.value and str(f.value).strip():
                            custom_h.append(float(f.value))
                        else:
                            custom_h.append(None)  # Mark as empty
                    except:
                        custom_h.append(None)
                
                # Get opening dimensions from input fields first
                try:
                    w_val = inputs["width"].value
                    opening_w = float(w_val) if w_val and str(w_val).strip() else 0
                except:
                    opening_w = 0
                
                try:
                    h_val = inputs["height"].value
                    opening_h = float(h_val) if h_val and str(h_val).strip() else 0
                except:
                    opening_h = 0
                
                # If opening dimensions not set, try to calculate from custom dimensions
                if opening_w <= 0:
                    filled_w = [w for w in custom_w if w is not None]
                    if len(filled_w) == bw:  # All widths filled
                        opening_w = sum(filled_w)
                    elif len(filled_w) > 0:  # Some widths filled, estimate total
                        avg_w = sum(filled_w) / len(filled_w)
                        opening_w = avg_w * bw
                
                if opening_h <= 0:
                    filled_h = [h for h in custom_h if h is not None]
                    if len(filled_h) == bh:  # All heights filled
                        opening_h = sum(filled_h)
                    elif len(filled_h) > 0:  # Some heights filled, estimate total
                        avg_h = sum(filled_h) / len(filled_h)
                        opening_h = avg_h * bh
                
                # If still no opening dimensions, calculate from bay counts (use reasonable defaults)
                if opening_w <= 0:
                    opening_w = max(bw * 36.0, 100.0)  # Default estimate, minimum 100
                if opening_h <= 0:
                    opening_h = max(bh * 48.0, 100.0)  # Default estimate, minimum 100
                
                # Prepare custom dimensions for diagram (only use if all are filled)
                final_custom_w = None
                if len(custom_w) == bw and all(w is not None for w in custom_w):
                    final_custom_w = [w for w in custom_w]
                
                final_custom_h = None
                if len(custom_h) == bh and all(h is not None for h in custom_h):
                    final_custom_h = [h for h in custom_h]
                
                # Generate diagram
                diagram_b64 = create_bay_diagram_base64(
                    bw, bh, opening_w, opening_h,
                    custom_bay_widths=final_custom_w,
                    custom_bay_heights=final_custom_h
                )
                
                if diagram_b64 and inputs.get("bay_diagram_image"):
                    inputs["bay_diagram_image"].src_base64 = diagram_b64
                    inputs["bay_diagram_image"].visible = True
                    if inputs.get("bay_diagram_container"):
                        inputs["bay_diagram_container"].visible = True
                    print(f"[Bay Diagram] Updated: {bw}W x {bh}H, opening={opening_w}x{opening_h}, custom_w={final_custom_w is not None}, custom_h={final_custom_h is not None}")
                    if page.views:
                        page.update()
                else:
                    # If diagram generation failed, hide the container
                    print(f"[Bay Diagram] Failed to generate: diagram_b64={diagram_b64 is not None}, image={inputs.get('bay_diagram_image') is not None}, bw={bw}, bh={bh}, opening_w={opening_w}, opening_h={opening_h}")
                    if inputs.get("bay_diagram_container"):
                        inputs["bay_diagram_container"].visible = False
                    if page.views:
                        page.update()
            except Exception as ex:
                # Log error but don't break the UI
                import traceback
                print(f"[Bay Diagram] Refresh error: {ex}")
                print(traceback.format_exc())
                if inputs.get("bay_diagram_container"):
                    inputs["bay_diagram_container"].visible = False
                if page.views:
                    page.update()

        def on_elevation_load(e):
            elev_name = inputs["saved_elev"].value
            if not elev_name or elev_name == "New Elevation": 
                # If "New Elevation" selected or cleared, reset to Create mode
                clear_workspace()
                inputs["save_btn"].text = "CREATE ELEVATION"
                inputs["saved_elev"].value = "New Elevation"  # Keep "New Elevation" for consistency
                # Unlock the name field and hide duplicate button in create mode
                inputs["type"].read_only = False
                if inputs.get("duplicate_btn"):
                    inputs["duplicate_btn"].visible = False
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
            # Lock the name field and show duplicate button in update mode
            inputs["type"].read_only = True
            if inputs.get("duplicate_btn"):
                inputs["duplicate_btn"].visible = True
            inputs["count"].value = str(data.get("total_count", ""))
            inputs["width"].value = str(data.get("opening_width_inches", ""))
            inputs["height"].value = str(data.get("opening_height_inches", ""))
            
            inputs["bays_wide"].value = str(data.get("bays_wide", ""))
            inputs["bays_tall"].value = str(data.get("bays_tall", ""))
            
            # Load and update quantity per elevation and cost per elevation checkboxes
            if inputs.get("qty_per_elev_checkbox"):
                inputs["qty_per_elev_checkbox"].value = data.get("show_qty_per_elevation", False)
            if inputs.get("total_cost_per_elev_checkbox"):
                inputs["total_cost_per_elev_checkbox"].value = data.get("show_total_cost_per_elevation", False)
            if inputs.get("discounted_cost_per_elev_checkbox"):
                inputs["discounted_cost_per_elev_checkbox"].value = data.get("show_discounted_cost_per_elevation", False)
            if inputs.get("door_only_checkbox"):
                inputs["door_only_checkbox"].value = data.get("door_only", False)
                update_door_only_visibility(None)  # Hide rest of form when door_only
            update_qty_per_elev_visibility(None)  # Update visibility based on count
            
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
            # Auto-refresh bay diagram after loading
            auto_refresh_bay_diagram()
            page.update()

        def clear_workspace():
            inputs["type"].value = ""
            inputs["type"].read_only = False  # Unlock name field in create mode
            inputs["count"].value = ""
            inputs["width"].value = ""
            inputs["height"].value = ""
            inputs["bays_wide"].value = ""
            inputs["bays_tall"].value = ""
            inputs["save_btn"].text = "CREATE ELEVATION"
            # Hide duplicate button in create mode
            if inputs.get("duplicate_btn"):
                inputs["duplicate_btn"].visible = False
            # Clear dynamic fields
            inputs["custom_w_col"].controls.clear()
            inputs["custom_h_col"].controls.clear()
            inputs["dynamic_w_fields"] = []
            inputs["dynamic_h_fields"] = []
            
            # Clear quantity per elevation and cost per elevation checkboxes
            if inputs.get("qty_per_elev_checkbox"):
                inputs["qty_per_elev_checkbox"].value = False
            if inputs.get("qty_per_elev_container"):
                inputs["qty_per_elev_container"].visible = False
            if inputs.get("total_cost_per_elev_checkbox"):
                inputs["total_cost_per_elev_checkbox"].value = False
            if inputs.get("discounted_cost_per_elev_checkbox"):
                inputs["discounted_cost_per_elev_checkbox"].value = False
            if inputs.get("cost_per_elev_container"):
                inputs["cost_per_elev_container"].visible = False
            if inputs.get("door_only_checkbox"):
                inputs["door_only_checkbox"].value = False
            update_door_only_visibility(None)  # Show full form again
            
            # Hide bay diagram preview
            if inputs.get("bay_diagram_container"):
                inputs["bay_diagram_container"].visible = False
            if inputs.get("bay_diagram_image"):
                inputs["bay_diagram_image"].visible = False
                inputs["bay_diagram_image"].src_base64 = ""
            
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
            
            # Validate door count - must be a positive integer
            door_count_str = inputs["door_count"].value.strip() if inputs["door_count"].value else ""
            if not door_count_str:
                show_snack("Door count is required")
                return
            try:
                count = int(door_count_str)
                if count <= 0:
                    show_snack("Door count must be greater than 0")
                    return
            except ValueError:
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
            
            # Clear door inputs - leave count blank
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
                # Auto-save miscellaneous cost settings before generating report
                existing_settings = load_project_settings(state["current_project"])
                misc_settings = {
                    "overhead_materials_pct": get_input_pct("overhead_materials_pct"),
                    "overhead_labor_pct": get_input_pct("overhead_labor_pct"),
                    "admin_management_pct": get_input_pct("admin_management_pct"),
                    "engineering_pct": get_input_pct("engineering_pct"),
                    "packaging_materials_pct": get_input_pct("packaging_materials_pct"),
                    "shipping_transport_pct": get_input_pct("shipping_transport_pct"),
                    "commissions_pct": get_input_pct("commissions_pct")
                }
                # Merge with existing settings to preserve markups
                existing_settings.update(misc_settings)
                
                paths = get_project_paths(state["current_project"])
                save_project_settings(state["current_project"], existing_settings)
                # Verify file was created and wait a moment for file system
                time.sleep(0.1)  # Small delay to ensure file is written
                
                elev = inputs["type"].value.strip() if inputs["type"].value else ""
                if not elev: raise ValueError("Elevation Name Required")
                
                door_only = inputs.get("door_only_checkbox") and inputs["door_only_checkbox"].value

                if door_only:
                    if not state["current_doors"]:
                        raise ValueError("Door only: At least one door is required")
                    total = 1
                    finish = state["finish_options"][0] if state["finish_options"] else "Clear"
                    system = "Other"
                    # Derive w, h from first door size (e.g. "3' X 7'" -> 36, 84)
                    first_door = state["current_doors"][0]
                    size_str = first_door.get("size", "3' X 7'")
                    m = re.search(r"(\d+)'\s*[xX]\s*(\d+)'", size_str)
                    if m:
                        w = float(m.group(1)) * 12
                        h = float(m.group(2)) * 12
                    else:
                        w, h = 36.0, 84.0
                else:
                    if not inputs["count"].value:
                        raise ValueError("Quantity is required")
                    total = int(inputs["count"].value)
                    if not inputs["width"].value:
                        raise ValueError("Opening Width is required")
                    w = float(inputs["width"].value)
                    if not inputs["height"].value:
                        raise ValueError("Opening Height is required")
                    h = float(inputs["height"].value)
                    finish = inputs["finish"].value
                    system = inputs["system"].value
                
                # Simple calculations for preview/saving
                sqft = calculate_rectangle_area(w/12, h/12)
                perim = calculate_perimeter(w/12, h/12)
                
                # Get the checkbox values
                show_qty_per_elev = inputs.get("qty_per_elev_checkbox", ft.Checkbox()).value if inputs.get("qty_per_elev_checkbox") else False
                show_total_cost_per_elev = inputs.get("total_cost_per_elev_checkbox", ft.Checkbox()).value if inputs.get("total_cost_per_elev_checkbox") else False
                show_discounted_cost_per_elev = inputs.get("discounted_cost_per_elev_checkbox", ft.Checkbox()).value if inputs.get("discounted_cost_per_elev_checkbox") else False

                data = {
                    "system": system,
                    "finish": finish,
                    "total_count": total,
                    "opening_width_inches": w,
                    "opening_height_inches": h,
                    "sqft_per_type": sqft,
                    "total_sqft": sqft * total,
                    "perimeter_ft": perim,
                    "total_perimeter_ft": perim * total,
                    "show_qty_per_elevation": show_qty_per_elev,
                    "show_total_cost_per_elevation": show_total_cost_per_elev,
                    "show_discounted_cost_per_elevation": show_discounted_cost_per_elev,
                    "door_only": door_only
                }

                if door_only:
                    if not state["current_doors"]:
                        raise ValueError("Door only: At least one door is required")
                    data["bays_wide"] = 0
                    data["bays_tall"] = 0
                    data["custom_bay_widths"] = []
                    data["custom_bay_heights"] = []
                    data["calculated_outputs"] = []
                elif data["system"] == "YES 45TU FRONT SET(OG)":
                    if not inputs["bays_wide"].value:
                        raise ValueError("Bays Wide is required for YES 45TU FRONT SET(OG)")
                    if not inputs["bays_tall"].value:
                        raise ValueError("Bays Tall is required for YES 45TU FRONT SET(OG)")
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
                        bw, bh, total, w, h, state["current_doors"], 
                        custom_bay_widths=w_vals if w_vals else None
                    )
                else:
                    data["calculated_outputs"] = []

                # Append formatted door items to calculated_outputs so they are persisted in state and file
                if state["current_doors"]:
                    door_items = calculate_door_info(state["current_doors"], finish=data["finish"], total_count=total)
                    data["calculated_outputs"].extend(door_items)

                state["saved_elevations"][elev] = data
                
                # Save to database
                try:
                    db.save_elevations(state["current_project"], state["saved_elevations"])
                except Exception as save_err:
                    error_msg = f"Failed to save elevations: {str(save_err)}"
                    print(f"[ERROR] {error_msg}")
                    show_snack(error_msg, "red")
                    raise
                
                # Reload elevations to ensure state is in sync
                load_elevations(state["current_project"])
                
                # Trigger Excel Gen (silent or with notif)
                try:
                    # Use in-memory elevations to ensure door-only and all data is included (avoid DB read lag)
                    paths = prepare_temp_files_for_excel(state["current_project"], elevations_override=state["saved_elevations"])
                    
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
                        custom_bay_heights=data.get("custom_bay_heights", []),
                        summary_settings_path=paths["settings"],
                        door_only=data.get("door_only", False)
                    )
                    
                    # Sync changes back to database and cleanup temp files
                    sync_from_temp_files(state["current_project"], paths)
                    
                    # Check if file was actually created
                    if os.path.exists(paths["excel"]):
                        show_snack(f"Report saved to: {paths['excel']}", "green")
                    else:
                        error_msg = f"Report file not found at: {paths['excel']}"
                        print(f"[ERROR] {error_msg}")
                        show_snack(error_msg, "red")
                except Exception as report_err:
                    error_msg = f"Report generation failed: {str(report_err)}"
                    print(f"[ERROR] {error_msg}")
                    show_snack(error_msg, "red")
                    traceback.print_exc()
                    # Cleanup temp files even on error
                    cleanup_temp_files(paths)
                
                # Update dropdown with fresh data
                new_opts = sorted(state["saved_elevations"].keys())
                dropdown_options = [ft.dropdown.Option("New Elevation")] + [ft.dropdown.Option(x) for x in new_opts]
                inputs["saved_elev"].options = dropdown_options
                
                if is_update:
                    inputs["saved_elev"].value = elev
                    show_snack("Elevation Updated Successfully", "green")
                else:
                    # Clear Inputs for new creation
                    clear_workspace()
                    inputs["saved_elev"].value = "New Elevation"  # Set to "New Elevation" instead of None
                    show_snack("Elevation Created Successfully", "green")
                
                # Force update the dropdown control
                inputs["saved_elev"].update()
                
                # Refresh waste calculator to reflect updated elevation data
                try:
                    refresh_waste_calculator()
                except Exception as ex:
                    print(f"[WARNING] Error refreshing waste calculator: {ex}")
                
                
                page.update()

            except Exception as ex:
                show_snack(f"Error: {str(ex)}", "red")

        def duplicate_elevation_action(e):
            """Duplicate the current elevation with a new name"""
            elev_name = inputs["saved_elev"].value
            if not elev_name or elev_name == "New Elevation" or elev_name not in state["saved_elevations"]:
                show_snack("No elevation selected to duplicate", "red")
                return
            
            # Get the current elevation data
            original_data = state["saved_elevations"][elev_name].copy()
            
            # Create a text field for new name
            new_name_field = ft.TextField(
                label="New Elevation Name",
                value=f"{elev_name} (Copy)",
                bgcolor=COLOR_INPUT_BG,
                border_color=COLOR_ACCENT_LIGHT,
                color=COLOR_TEXT,
                label_style=ft.TextStyle(color=COLOR_TEXT_DIM),
                focused_border_color=COLOR_ACCENT,
                autofocus=True,
            )
            
            def do_duplicate(e):
                new_name = new_name_field.value.strip()
                if not new_name:
                    show_snack("Elevation name cannot be empty", "red")
                    return
                
                if new_name in state["saved_elevations"]:
                    show_snack(f"Elevation '{new_name}' already exists", "red")
                    return
                
                # Create duplicate with new name
                state["saved_elevations"][new_name] = original_data.copy()
                
                # Copy door data if it exists
                try:
                    door_data = db.get_doors(state["current_project"], elev_name)
                    if door_data:
                        db.save_doors(state["current_project"], new_name, door_data)
                except Exception as ex:
                    print(f"Error copying door data: {ex}")
                
                # Save elevations to database
                db.save_elevations(state["current_project"], state["saved_elevations"])
                
                # Reload elevations
                load_elevations(state["current_project"])
                
                # Update dropdown
                new_opts = sorted(state["saved_elevations"].keys())
                dropdown_options = [ft.dropdown.Option("New Elevation")] + [ft.dropdown.Option(x) for x in new_opts]
                inputs["saved_elev"].options = dropdown_options
                inputs["saved_elev"].value = new_name
                
                # Close dialog
                dlg.open = False
                page.update()
                
                # Load the duplicated elevation
                on_elevation_load(None)
                
                show_snack(f"Elevation duplicated as '{new_name}'", "green")
            
            def cancel_duplicate(e):
                dlg.open = False
                page.update()
            
            dlg = ft.AlertDialog(
                modal=True,
                title=ft.Text("Duplicate Elevation", color=COLOR_TEXT),
                content=ft.Container(
                    content=new_name_field,
                    width=400,
                    padding=10
                ),
                actions=[
                    ft.TextButton("Cancel", on_click=cancel_duplicate),
                    ft.TextButton("Duplicate", on_click=do_duplicate, style=ft.ButtonStyle(color=COLOR_ACCENT))
                ],
                bgcolor=COLOR_SURFACE,
            )
            
            page.overlay.append(dlg)
            dlg.open = True
            page.update()

        def delete_elevation_action(e):
            elev = inputs["saved_elev"].value
            if elev and elev != "New Elevation" and elev in state["saved_elevations"]:
                # Delete elevation and its doors from database
                db.delete_elevation(state["current_project"], elev)
                
                # Reload elevations to ensure state is in sync
                load_elevations(state["current_project"])
                
                # Regenerate the report to reflect deletion
                try:
                    temp_paths = prepare_temp_files_for_excel(state["current_project"])
                    generate_excel_report(
                        excel_path=temp_paths["excel"], 
                        elevations_json_path=temp_paths["elevations"], 
                        extra_materials_json_path=temp_paths["materials"],
                        system_input="", finish_input="", elevation_type="", total_count=0,
                        bays_wide=0, bays_tall=0, opening_width=0, opening_height=0,
                        sqft_per_type=0, total_sqft=0, perimeter_ft=0, total_perimeter_ft=0,
                        calculated_outputs=[], doors=None, 
                        delete_elevation_type=elev,
                        summary_settings_path=temp_paths["settings"]
                    )
                    sync_from_temp_files(state["current_project"], temp_paths)
                except Exception as ex:
                    print(f"Error regenerating report after delete: {ex}")
                    cleanup_temp_files(temp_paths)

                # Update UI with fresh data
                new_opts = sorted(state["saved_elevations"].keys())
                dropdown_options = [ft.dropdown.Option("New Elevation")] + [ft.dropdown.Option(x) for x in new_opts]
                inputs["saved_elev"].options = dropdown_options
                inputs["saved_elev"].value = "New Elevation"
                # Force switch to CREATE mode after deletion by clearing inputs
                clear_workspace()
                # Since clear_workspace resets the button text, we just need to confirm
                
                # Refresh waste calculator to reflect deleted elevation data
                try:
                    refresh_waste_calculator()
                except Exception as ex:
                    print(f"[WARNING] Error refreshing waste calculator: {ex}")
                
                show_snack("Elevation Deleted", "red")
                # Force update the dropdown control
                inputs["saved_elev"].update()
                page.update()
            else:
                show_snack("Please select an elevation to delete", "red")

        def gen_full_report(e):
            temp_paths = None
            try:
                settings = {k: get_input_pct(k) for k in [
                    "overhead_materials_pct", "overhead_labor_pct", "admin_management_pct",
                    "engineering_pct", "packaging_materials_pct", "shipping_transport_pct", "commissions_pct"
                ]}
                save_project_settings(state["current_project"], settings)
                
                temp_paths = prepare_temp_files_for_excel(state["current_project"])
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                out = os.path.join("reports", f"{state['current_project']}_{ts}.xlsx")
                os.makedirs("reports", exist_ok=True)
                
                generate_excel_report(
                    out, temp_paths["elevations"], temp_paths["materials"],
                    "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, [], None, None, mode="export_all",
                    summary_settings_path=temp_paths["settings"]
                )
                sync_from_temp_files(state["current_project"], temp_paths)
                show_snack(f"Full Report Generated: {out}", "green")
            except Exception as ex:
                error_msg = f"Report Error: {ex}"
                print(f"[ERROR] {error_msg}")
                traceback.print_exc()
                show_snack(error_msg, "red")
                if temp_paths:
                    cleanup_temp_files(temp_paths)
        
        def export_to_pdf(e):
            """Export current project to PDF"""
            temp_paths = None
            try:
                if not REPORTLAB_AVAILABLE:
                    show_snack("PDF export requires reportlab. Install with: pip install reportlab", "red")
                    return
                
                if not state.get("current_project"):
                    show_snack("Please select a project first", "red")
                    return
                
                # Find the most recent Excel report or generate one
                paths = get_project_paths(state["current_project"])
                excel_path = paths.get("excel")
                
                # If no report exists, generate one first
                if not excel_path or not os.path.exists(excel_path):
                    show_snack("Generating report first...", "blue")
                    # Generate report
                    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    excel_path = os.path.join("reports", f"{state['current_project']}_{ts}.xlsx")
                    os.makedirs("reports", exist_ok=True)
                    
                    # Save settings first
                    settings = {k: get_input_pct(k) for k in [
                        "overhead_materials_pct", "overhead_labor_pct", "admin_management_pct",
                        "engineering_pct", "packaging_materials_pct", "shipping_transport_pct", "commissions_pct"
                    ]}
                    save_project_settings(state["current_project"], settings)
                    
                    temp_paths = prepare_temp_files_for_excel(state["current_project"])
                    generate_excel_report(
                        excel_path, temp_paths["elevations"], temp_paths["materials"],
                        "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, [], None, None, mode="export_all",
                        summary_settings_path=temp_paths["settings"]
                    )
                    sync_from_temp_files(state["current_project"], temp_paths)
                
                # Export to PDF
                pdf_path = export_project_to_pdf(
                    state["current_project"],
                    excel_path=excel_path,
                    include_logo=True
                )
                
                if os.path.exists(pdf_path):
                    show_snack(f"PDF exported successfully: {pdf_path}", "green")
                    # Open the PDF (Windows)
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                else:
                    show_snack("PDF export failed", "red")
                    
            except Exception as ex:
                error_msg = f"PDF export error: {str(ex)}"
                print(f"[ERROR] {error_msg}")
                traceback.print_exc()
                show_snack(error_msg, "red")
                if temp_paths:
                    cleanup_temp_files(temp_paths)
        
        # --- UI Structure ---
        
        def save_summary_settings(e):
            temp_paths = None
            try:
                settings = {k: get_input_pct(k) for k in [
                    "overhead_materials_pct", "overhead_labor_pct", "admin_management_pct",
                    "engineering_pct", "packaging_materials_pct", "shipping_transport_pct", "commissions_pct"
                ]}
                
                # Load existing settings and merge with miscellaneous settings to preserve markups
                existing_settings = load_project_settings(state["current_project"])
                existing_settings.update(settings)
                
                # Save settings to database
                save_project_settings(state["current_project"], existing_settings)
                
                # Check if any percentages are > 0
                has_percentages = any(pct > 0 for pct in settings.values())
                
                # Regenerate the Excel report to show the updated miscellaneous summary
                if state["saved_elevations"]:
                    try:
                        # Generate to a temporary location first, then copy to main project file
                        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        os.makedirs("reports", exist_ok=True)
                        temp_report_path = os.path.join("reports", f"{state['current_project']}_temp_{ts}.xlsx")
                        
                        temp_paths = prepare_temp_files_for_excel(state["current_project"])
                        generate_excel_report(
                            temp_report_path, 
                            temp_paths["elevations"], 
                            temp_paths["materials"],
                            "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, [], 
                            None, False, None, None, mode="export_all",
                            summary_settings_path=temp_paths["settings"]
                        )
                        sync_from_temp_files(state["current_project"], temp_paths)
                        
                        # Copy the generated report to the main project Excel file location
                        paths = get_project_paths(state["current_project"])
                        if os.path.exists(temp_report_path):
                            # Ensure the directory exists
                            excel_dir = os.path.dirname(paths["excel"])
                            if excel_dir:
                                os.makedirs(excel_dir, exist_ok=True)
                            
                            # Copy to main project file
                            shutil.copy2(temp_report_path, paths["excel"])
                            
                            # Clean up temp file
                            try:
                                os.remove(temp_report_path)
                            except:
                                pass
                        if has_percentages:
                            show_snack("Miscellaneous cost settings saved and report updated. Please close and reopen Excel to see changes.", "green")
                        else:
                            show_snack("Settings saved (all percentages are 0%)", "orange")
                    except Exception as report_err:
                        error_msg = f"Settings saved but report update failed: {str(report_err)}"
                        print(f"[ERROR] {error_msg}")
                        traceback.print_exc()
                        show_snack("Settings saved, but report update failed", "orange")
                        if temp_paths:
                            cleanup_temp_files(temp_paths)
                else:
                    if has_percentages:
                        show_snack("Miscellaneous cost settings saved", "green")
                    else:
                        show_snack("Settings saved (all percentages are 0%)", "orange")
                
                page.update()
            except Exception as ex:
                error_msg = f"Error saving settings: {str(ex)}"
                print(f"[ERROR] {error_msg}")
                if temp_paths:
                    cleanup_temp_files(temp_paths)
                traceback.print_exc()
                show_snack(error_msg, "red")
        
        def save_markup_settings(e):
            temp_paths = None
            try:
                settings = {k: get_input_pct(k) for k in [
                    "profit_on_material_pct", "profit_on_waste_pct", "profit_on_glass_pct",
                    "profit_on_wages_pct", "planning_technical_pct", "commission_pct"
                ]}
                
                # Load existing settings and merge with markups
                existing_settings = load_project_settings(state["current_project"])
                existing_settings.update(settings)
                
                # Save settings to database
                save_project_settings(state["current_project"], existing_settings)
                
                # Check if any percentages are > 0
                has_percentages = any(pct > 0 for pct in settings.values())
                
                # Regenerate the Excel report to show the updated markups
                if state["saved_elevations"]:
                    try:
                        # Generate to a temporary location first, then copy to main project file
                        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        os.makedirs("reports", exist_ok=True)
                        temp_report_path = os.path.join("reports", f"{state['current_project']}_temp_{ts}.xlsx")
                        
                        temp_paths = prepare_temp_files_for_excel(state["current_project"])
                        generate_excel_report(
                            temp_report_path, 
                            temp_paths["elevations"], 
                            temp_paths["materials"],
                            "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, [], 
                            None, False, None, None, mode="export_all",
                            summary_settings_path=temp_paths["settings"]
                        )
                        sync_from_temp_files(state["current_project"], temp_paths)
                        
                        # Copy the generated report to the main project Excel file location
                        paths = get_project_paths(state["current_project"])
                        if os.path.exists(temp_report_path):
                            # Ensure the directory exists
                            excel_dir = os.path.dirname(paths["excel"])
                            if excel_dir:
                                os.makedirs(excel_dir, exist_ok=True)
                            
                            # Copy to main project file
                            shutil.copy2(temp_report_path, paths["excel"])
                            
                            # Clean up temp file
                            try:
                                os.remove(temp_report_path)
                            except:
                                pass
                        if has_percentages:
                            show_snack("Markup settings saved and report updated. Please close and reopen Excel to see changes.", "green")
                        else:
                            show_snack("Settings saved (all percentages are 0%)", "orange")
                    except Exception as report_err:
                        error_msg = f"Settings saved but report update failed: {str(report_err)}"
                        print(f"[ERROR] {error_msg}")
                        traceback.print_exc()
                        show_snack("Settings saved, but report update failed", "orange")
                        if temp_paths:
                            cleanup_temp_files(temp_paths)
                else:
                    if has_percentages:
                        show_snack("Markup settings saved", "green")
                    else:
                        show_snack("Settings saved (all percentages are 0%)", "orange")
                
                page.update()
            except Exception as ex:
                error_msg = f"Error saving markup settings: {str(ex)}"
                print(f"[ERROR] {error_msg}")
                traceback.print_exc()
                show_snack(error_msg, "red")
                if temp_paths:
                    cleanup_temp_files(temp_paths)
        
        def save_elevation_summary_settings(e):
            temp_paths = None
            try:
                settings = {
                    "show_elevation_names": inputs.get("show_elevation_names", ft.Checkbox()).value if inputs.get("show_elevation_names") else False,
                    "show_elevation_quantity": inputs.get("show_elevation_quantity", ft.Checkbox()).value if inputs.get("show_elevation_quantity") else False,
                    "show_elevation_dimensions": inputs.get("show_elevation_dimensions", ft.Checkbox()).value if inputs.get("show_elevation_dimensions") else False,
                    "show_elevation_sqft": inputs.get("show_elevation_sqft", ft.Checkbox()).value if inputs.get("show_elevation_sqft") else False,
                    "show_elevation_perimeter": inputs.get("show_elevation_perimeter", ft.Checkbox()).value if inputs.get("show_elevation_perimeter") else False
                }
                
                existing_settings = load_project_settings(state["current_project"])
                existing_settings.update(settings)
                save_project_settings(state["current_project"], existing_settings)
                
                # Regenerate report if elevations exist
                if state["saved_elevations"]:
                    try:
                        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        os.makedirs("reports", exist_ok=True)
                        temp_report_path = os.path.join("reports", f"{state['current_project']}_temp_{ts}.xlsx")
                        
                        temp_paths = prepare_temp_files_for_excel(state["current_project"])
                        generate_excel_report(
                            temp_report_path, 
                            temp_paths["elevations"], 
                            temp_paths["materials"],
                            "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, [], 
                            None, False, None, None, mode="export_all",
                            summary_settings_path=temp_paths["settings"]
                        )
                        sync_from_temp_files(state["current_project"], temp_paths)
                        
                        paths = get_project_paths(state["current_project"])
                        if os.path.exists(temp_report_path):
                            excel_dir = os.path.dirname(paths["excel"])
                            if excel_dir:
                                os.makedirs(excel_dir, exist_ok=True)
                            shutil.copy2(temp_report_path, paths["excel"])
                            try:
                                os.remove(temp_report_path)
                            except:
                                pass
                        show_snack("Elevation summary settings saved and report updated", "green")
                    except Exception as report_err:
                        show_snack("Settings saved, but report update failed", "orange")
                        if temp_paths:
                            cleanup_temp_files(temp_paths)
                else:
                    show_snack("Elevation summary settings saved", "green")
                
                page.update()
            except Exception as ex:
                error_msg = f"Error saving settings: {str(ex)}"
                print(f"[ERROR] {error_msg}")
                traceback.print_exc()
                show_snack(error_msg, "red")
                if temp_paths:
                    cleanup_temp_files(temp_paths)
        
        # Combined Summary Section (Miscellaneous Cost + Markups)
        # Load saved settings first
        project_settings = load_project_settings(state["current_project"])
        
        # Miscellaneous Cost Section
        misc_section = ft.Column([
            ft.Row([
                ft.Text("MISCELLANEOUS COST SETTINGS", size=16, weight="bold", color=COLOR_ACCENT, expand=True),
                ft.ElevatedButton(
                    "SAVE", 
                    bgcolor=COLOR_ACCENT, 
                    color="white", 
                    on_click=save_summary_settings,
                    height=40,
                    width=120
                )
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            ft.Text("Enter percentages for summary calculations (leave blank for 0%)", size=12, color=COLOR_TEXT_DIM, italic=True),
            ft.Divider(height=1, color=COLOR_SURFACE),
            ft.Container(height=20),  # Spacing
            ft.Row([
                ft.Column([
                    create_input_field("Overhead Materials %", "overhead_materials_pct", numeric=True, expand=False),
                    create_input_field("Admin and Management %", "admin_management_pct", numeric=True, expand=False),
                    create_input_field("Packaging Materials %", "packaging_materials_pct", numeric=True, expand=False),
                    create_input_field("Commissions %", "commissions_pct", numeric=True, expand=False)
                ], spacing=15, expand=True),
                ft.Container(width=30),  # Spacing between columns
                ft.Column([
                    create_input_field("Overhead Labor %", "overhead_labor_pct", numeric=True, expand=False),
                    create_input_field("Engineering %", "engineering_pct", numeric=True, expand=False),
                    create_input_field("Shipping and Transport %", "shipping_transport_pct", numeric=True, expand=False),
                    ft.Container(height=48)  # Spacer to align with 4th field
                ], spacing=15, expand=True)
            ], spacing=0)
        ], spacing=15)
        
        # Markups Section
        markup_section = ft.Column([
            ft.Row([
                ft.Text("MARKUP SETTINGS", size=16, weight="bold", color=COLOR_ACCENT, expand=True),
                ft.ElevatedButton(
                    "SAVE", 
                    bgcolor=COLOR_ACCENT, 
                    color="white", 
                    on_click=save_markup_settings,
                    height=40,
                    width=120
                )
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            ft.Text("Enter percentages for markup calculations (leave blank for 0%)", size=12, color=COLOR_TEXT_DIM, italic=True),
            ft.Divider(height=1, color=COLOR_SURFACE),
            ft.Container(height=20),  # Spacing
            ft.Row([
                ft.Column([
                    create_input_field("Profit on Material %", "profit_on_material_pct", numeric=True, expand=False),
                    create_input_field("Profit on Glass Purchase %", "profit_on_glass_pct", numeric=True, expand=False),
                    create_input_field("Planning / Technical Office %", "planning_technical_pct", numeric=True, expand=False),
                    ft.Container(height=48)  # Spacer
                ], spacing=15, expand=True),
                ft.Container(width=30),  # Spacing between columns
                ft.Column([
                    create_input_field("Profit on Waste %", "profit_on_waste_pct", numeric=True, expand=False),
                    create_input_field("Profit on Wages %", "profit_on_wages_pct", numeric=True, expand=False),
                    create_input_field("Commission %", "commission_pct", numeric=True, expand=False),
                    ft.Container(height=48)  # Spacer
                ], spacing=15, expand=True)
            ], spacing=0)
        ], spacing=15)
        
        # Elevation Summary Display Options Section
        elevation_name_cb = ft.Checkbox(label="Elevation Names", value=False, fill_color=COLOR_ACCENT)
        elevation_quantity_cb = ft.Checkbox(label="Quantity", value=False, fill_color=COLOR_ACCENT)
        elevation_dimensions_cb = ft.Checkbox(label="Dimensions", value=False, fill_color=COLOR_ACCENT)
        elevation_sqft_cb = ft.Checkbox(label="SQFT Total", value=False, fill_color=COLOR_ACCENT)
        elevation_perimeter_cb = ft.Checkbox(label="Perimeter FT Total", value=False, fill_color=COLOR_ACCENT)
        
        inputs['show_elevation_names'] = elevation_name_cb
        inputs['show_elevation_quantity'] = elevation_quantity_cb
        inputs['show_elevation_dimensions'] = elevation_dimensions_cb
        inputs['show_elevation_sqft'] = elevation_sqft_cb
        inputs['show_elevation_perimeter'] = elevation_perimeter_cb
        
        elevation_summary_section = ft.Column([
            ft.Row([
                ft.Text("ELEVATION SUMMARY DISPLAY", size=16, weight="bold", color=COLOR_ACCENT, expand=True),
                ft.ElevatedButton(
                    "SAVE", 
                    bgcolor=COLOR_ACCENT, 
                    color="white", 
                    on_click=save_elevation_summary_settings,
                    height=40,
                    width=120
                )
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            ft.Text("Select which columns to display in the summary report", size=12, color=COLOR_TEXT_DIM, italic=True),
            ft.Divider(height=1, color=COLOR_SURFACE),
            ft.Container(height=20),  # Spacing
            ft.Row([
                ft.Column([
                    elevation_name_cb,
                    elevation_quantity_cb,
                    elevation_dimensions_cb
                ], spacing=10, expand=True),
                ft.Container(width=30),  # Spacing between columns
                ft.Column([
                    elevation_sqft_cb,
                    elevation_perimeter_cb,
                    ft.Container(height=48)  # Spacer
                ], spacing=10, expand=True)
            ], spacing=0)
        ], spacing=15)
        
        # Load saved elevation summary display settings
        inputs["show_elevation_names"].value = project_settings.get("show_elevation_names", False)
        inputs["show_elevation_quantity"].value = project_settings.get("show_elevation_quantity", False)
        inputs["show_elevation_dimensions"].value = project_settings.get("show_elevation_dimensions", False)
        inputs["show_elevation_sqft"].value = project_settings.get("show_elevation_sqft", False)
        inputs["show_elevation_perimeter"].value = project_settings.get("show_elevation_perimeter", False)
        
        # Waste Calculator Section
        def refresh_waste_calculator():
            """Refresh waste calculator display with current project data."""
            if not state.get("current_project"):
                waste_stats = {
                    "total_waste_cost": 0.0,
                    "total_material_cost": 0.0,
                    "overall_waste_percentage": 0.0,
                    "material_breakdown": [],
                    "suggestions": []
                }
            else:
                # Get data directly from Supabase database
                elevations_data = db.get_elevations(state["current_project"])
                extra_materials_data = db.get_materials(state["current_project"])
                
                # Debug output
                print(f"[Waste Calculator Frontend] Elevations: {len(elevations_data)} items")
                print(f"[Waste Calculator Frontend] Extra materials: {len(extra_materials_data)} items")
                if extra_materials_data:
                    print(f"[Waste Calculator Frontend] Extra materials keys: {list(extra_materials_data.keys())[:5]}")
                
                # Get Excel path if it exists
                excel_path = get_project_paths(state["current_project"]).get("excel")
                if excel_path and not os.path.exists(excel_path):
                    excel_path = None
                
                # Calculate waste statistics directly from database data
                try:
                    waste_stats = calculate_waste_statistics(
                        elevations_data=elevations_data,
                        extra_materials=extra_materials_data,
                        excel_path=excel_path
                    )
                    print(f"[Waste Calculator Frontend] Material breakdown: {len(waste_stats.get('material_breakdown', []))} items")
                except Exception as e:
                    print(f"[ERROR] Waste calculator error: {e}")
                    import traceback
                    traceback.print_exc()
                    waste_stats = {
                        "total_waste_cost": 0.0,
                        "total_material_cost": 0.0,
                        "overall_waste_percentage": 0.0,
                        "material_breakdown": [],
                        "suggestions": []
                    }
            
            # Update waste percentage display
            waste_pct = waste_stats["overall_waste_percentage"]
            waste_percentage_text.value = f"{waste_pct:.2f}%"
            waste_percentage_text.color = get_waste_percentage_color(waste_pct)
            
            # Update progress bar
            waste_progress.value = min(waste_pct / 100.0, 1.0)
            waste_progress.color = get_waste_percentage_color(waste_pct)
            
            # Update cost display
            total_waste_cost_text.value = f"${waste_stats['total_waste_cost']:.2f}"
            total_material_cost_text.value = f"${waste_stats['total_material_cost']:.2f}"
            
            # Update material breakdown table
            breakdown_rows = []
            for material in waste_stats["material_breakdown"]:
                breakdown_rows.append(
                    ft.DataRow(
                        cells=[
                            ft.DataCell(ft.Text(material["description"], size=11)),
                            ft.DataCell(ft.Text(f"{material['waste_percentage']:.1f}%", size=11, 
                                              color=get_waste_percentage_color(material["waste_percentage"]))),
                            ft.DataCell(ft.Text(f"${material['waste_cost']:.2f}", size=11)),
                            ft.DataCell(ft.Text(
                                material.get('waste_quantity_display', f"{material['waste_quantity']:.2f} {material['unit']}"), 
                                size=11
                            ))
                        ]
                    )
                )
            
            if not breakdown_rows:
                breakdown_rows.append(
                    ft.DataRow(
                        cells=[
                            ft.DataCell(ft.Text("No waste data available", italic=True, size=11, color=COLOR_TEXT_DIM)),
                            ft.DataCell(ft.Text("")),
                            ft.DataCell(ft.Text("")),
                            ft.DataCell(ft.Text(""))
                        ]
                    )
                )
            
            waste_breakdown_table.rows = breakdown_rows
            
            # Update suggestions
            suggestions_list.controls = []
            for suggestion in waste_stats["suggestions"]:
                priority_color = {
                    "high": "#F44336",  # Red
                    "medium": "#FFC107",  # Yellow
                    "low": "#4CAF50"  # Green
                }.get(suggestion["priority"], COLOR_TEXT)
                
                suggestions_list.controls.append(
                    ft.Container(
                        content=ft.Row([
                            ft.Container(
                                width=8,
                                height=8,
                                bgcolor=priority_color,
                                border_radius=4,
                                margin=ft.margin.only(right=10, top=6)
                            ),
                            ft.Text(suggestion["message"], size=11, color=COLOR_TEXT, expand=True)
                        ]),
                        padding=10,
                        bgcolor=COLOR_INPUT_BG,
                        border_radius=5,
                        margin=ft.margin.only(bottom=8)
                    )
                )
            
            if not suggestions_list.controls:
                suggestions_list.controls.append(
                    ft.Container(
                        content=ft.Text("No suggestions available", italic=True, size=11, color=COLOR_TEXT_DIM),
                        padding=10
                    )
                )
            
            page.update()
        
        # Waste calculator UI components
        waste_percentage_text = ft.Text("0.00%", size=24, weight="bold", color=COLOR_ACCENT)
        waste_progress = ft.ProgressBar(value=0.0, width=400, height=20, color=COLOR_ACCENT)
        total_waste_cost_text = ft.Text("$0.00", size=16, color=COLOR_TEXT)
        total_material_cost_text = ft.Text("$0.00", size=16, color=COLOR_TEXT)
        
        waste_breakdown_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Material", size=11, weight="bold", color=COLOR_ACCENT)),
                ft.DataColumn(ft.Text("Waste %", size=11, weight="bold", color=COLOR_ACCENT)),
                ft.DataColumn(ft.Text("Waste Cost", size=11, weight="bold", color=COLOR_ACCENT)),
                ft.DataColumn(ft.Text("Waste Qty", size=11, weight="bold", color=COLOR_ACCENT))
            ],
            rows=[],
            heading_row_color=COLOR_SURFACE,
            data_row_color={i: COLOR_INPUT_BG for i in range(100)},
            border=ft.Border(
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                ft.BorderSide(1, COLOR_ACCENT_LIGHT)
            )
        )
        
        suggestions_list = ft.Column([], spacing=0, scroll=ft.ScrollMode.AUTO, expand=True)
        
        waste_calculator_section = ft.Column([
            ft.Row([
                ft.Text("WASTE CALCULATOR", size=16, weight="bold", color=COLOR_ACCENT, expand=True),
                ft.ElevatedButton(
                    "REFRESH",
                    bgcolor=COLOR_ACCENT,
                    color="white",
                    on_click=lambda e: refresh_waste_calculator(),
                    height=40,
                    width=120
                )
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            ft.Text("Visual waste analysis and optimization suggestions", size=12, color=COLOR_TEXT_DIM, italic=True),
            ft.Divider(height=1, color=COLOR_SURFACE),
            ft.Container(height=20),
            
            # Visual waste percentage impact
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.ANALYTICS, color=COLOR_ACCENT, size=20),
                        ft.Text("Overall Waste Percentage", size=14, weight="bold", color=COLOR_TEXT, expand=True)
                    ]),
                    ft.Container(height=15),
                    ft.Row([
                        waste_percentage_text,
                        ft.Container(width=20),
                        ft.Container(
                            content=waste_progress,
                            expand=True
                        )
                    ], alignment=ft.MainAxisAlignment.START),
                    ft.Container(height=20),
                    ft.Row([
                        ft.Container(
                            content=ft.Column([
                                ft.Text("Total Waste Cost", size=11, color=COLOR_TEXT_DIM, weight="bold"),
                                total_waste_cost_text
                            ], spacing=5),
                            padding=15,
                            bgcolor=COLOR_INPUT_BG,
                            border_radius=8,
                            expand=True
                        ),
                        ft.Container(width=15),
                        ft.Container(
                            content=ft.Column([
                                ft.Text("Total Material Cost", size=11, color=COLOR_TEXT_DIM, weight="bold"),
                                total_material_cost_text
                            ], spacing=5),
                            padding=15,
                            bgcolor=COLOR_INPUT_BG,
                            border_radius=8,
                            expand=True
                        )
                    ])
                ]),
                padding=25,
                bgcolor=COLOR_SURFACE,
                border=ft.Border(
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT)
                ),
                border_radius=12,
                margin=ft.margin.only(bottom=20)
            ),
            
            # Waste breakdown by material type
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.TABLE_CHART, color=COLOR_ACCENT, size=20),
                        ft.Text("Waste Breakdown by Material Type", size=14, weight="bold", color=COLOR_TEXT, expand=True)
                    ]),
                    ft.Container(height=15),
                    ft.Container(
                        content=ft.Column([
                            waste_breakdown_table
                        ], scroll=ft.ScrollMode.AUTO, expand=True),
                        height=300,
                        border=ft.Border(
                            ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                            ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                            ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                            ft.BorderSide(1, COLOR_ACCENT_LIGHT)
                        ),
                        border_radius=8,
                        padding=5
                    )
                ]),
                padding=25,
                bgcolor=COLOR_SURFACE,
                border=ft.Border(
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT)
                ),
                border_radius=12,
                margin=ft.margin.only(bottom=20)
            ),
            
            # Optimization suggestions
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.LIGHTBULB, color=COLOR_ACCENT, size=20),
                        ft.Text("Optimization Suggestions", size=14, weight="bold", color=COLOR_TEXT, expand=True)
                    ]),
                    ft.Container(height=15),
                    ft.Container(
                        content=suggestions_list,
                        height=200,
                        border=ft.Border(
                            ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                            ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                            ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                            ft.BorderSide(1, COLOR_ACCENT_LIGHT)
                        ),
                        border_radius=8,
                        padding=5
                    )
                ]),
                padding=25,
                bgcolor=COLOR_SURFACE,
                border=ft.Border(
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT),
                    ft.BorderSide(1, COLOR_ACCENT_LIGHT)
                ),
                border_radius=12
            )
        ], spacing=15)
        
        # Combined Summary Container (scrollable)
        summary_settings_container = ft.Column([
            misc_section,
            ft.Container(height=30),  # Spacing between sections
            markup_section,
            ft.Container(height=30),  # Spacing between sections
            elevation_summary_section,
            ft.Container(height=30),  # Spacing between sections
            waste_calculator_section
        ], spacing=0, scroll=ft.ScrollMode.AUTO, expand=True)
        
        # Initialize waste calculator when project is loaded
        refresh_waste_calculator()
        
        # Load saved summary percentages after fields are created
        if inputs.get("overhead_materials_pct"):
            pct = project_settings.get("overhead_materials_pct", 0.0)
            inputs["overhead_materials_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("overhead_labor_pct"):
            pct = project_settings.get("overhead_labor_pct", 0.0)
            inputs["overhead_labor_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("admin_management_pct"):
            pct = project_settings.get("admin_management_pct", 0.0)
            inputs["admin_management_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("engineering_pct"):
            pct = project_settings.get("engineering_pct", 0.0)
            inputs["engineering_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("packaging_materials_pct"):
            pct = project_settings.get("packaging_materials_pct", 0.0)
            inputs["packaging_materials_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("shipping_transport_pct"):
            pct = project_settings.get("shipping_transport_pct", 0.0)
            inputs["shipping_transport_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("commissions_pct"):
            pct = project_settings.get("commissions_pct", 0.0)
            inputs["commissions_pct"].value = str(pct) if pct > 0 else ""
        
        # Load saved markup percentages after fields are created
        if inputs.get("profit_on_material_pct"):
            pct = project_settings.get("profit_on_material_pct", 0.0)
            inputs["profit_on_material_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("profit_on_waste_pct"):
            pct = project_settings.get("profit_on_waste_pct", 0.0)
            inputs["profit_on_waste_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("profit_on_glass_pct"):
            pct = project_settings.get("profit_on_glass_pct", 0.0)
            inputs["profit_on_glass_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("profit_on_wages_pct"):
            pct = project_settings.get("profit_on_wages_pct", 0.0)
            inputs["profit_on_wages_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("planning_technical_pct"):
            pct = project_settings.get("planning_technical_pct", 0.0)
            inputs["planning_technical_pct"].value = str(pct) if pct > 0 else ""
        if inputs.get("commission_pct"):
            pct = project_settings.get("commission_pct", 0.0)
            inputs["commission_pct"].value = str(pct) if pct > 0 else ""
        
        # Header
        header = ft.Row([
            ft.IconButton(ft.Icons.ARROW_BACK, on_click=lambda e: page.go("/"), icon_color=COLOR_TEXT),
            ft.Text(state["current_project"].upper(), size=20, weight="bold", color=COLOR_TEXT),
            ft.Container(expand=True),
            ft.Row([
                ft.ElevatedButton(
                    "EXPORT PDF",
                    icon=ft.Icons.PICTURE_AS_PDF,
                    bgcolor="#D32F2F",
                    color="white",
                    on_click=export_to_pdf,
                    tooltip="Export report as branded PDF with company logo"
                ) if REPORTLAB_AVAILABLE else ft.Container(),
                ft.ElevatedButton("GENERATE REPORT", bgcolor=COLOR_ACCENT, color="white", on_click=gen_full_report)
            ], spacing=10)
        ], height=60, alignment=ft.MainAxisAlignment.START)

        # Left Col: Elevation Form
        elev_options = sorted(state["saved_elevations"].keys())
        # Add "New Elevation" option at the top
        dd_options = ["New Elevation"] + elev_options
        saved_dd = create_dropdown("Load Elevation", "saved_elev", dd_options, on_change=on_elevation_load)
        
        # Form Group
        # Create quantity per elevation checkbox first (outside the Column list)
        qty_per_elev_checkbox = ft.Checkbox(
            label="Show 'Quantity Per Elevation' column in report",
            value=False,
            fill_color=COLOR_ACCENT
        )
        inputs["qty_per_elev_checkbox"] = qty_per_elev_checkbox
        
        # Create cost per elevation checkboxes
        total_cost_per_elev_checkbox = ft.Checkbox(
            label="Show 'Total List Cost Per Elevation' column in report",
            value=False,
            fill_color=COLOR_ACCENT
        )
        inputs["total_cost_per_elev_checkbox"] = total_cost_per_elev_checkbox
        
        discounted_cost_per_elev_checkbox = ft.Checkbox(
            label="Show 'Discounted Total List Cost Per Elevation' column in report",
            value=False,
            fill_color=COLOR_ACCENT
        )
        inputs["discounted_cost_per_elev_checkbox"] = discounted_cost_per_elev_checkbox
        
        # Door Manager (Collapsible) - defined before form_col so it can be referenced
        hardware_cbs = {opt: ft.Checkbox(label=opt, fill_color=COLOR_ACCENT) for opt in state["hardware_options"]}
        
        door_list_col = ft.Column(scroll=ft.ScrollMode.AUTO, spacing=5)
        
        door_manager_content = ft.Column([
            ft.Row([create_dropdown("Size", "door_size", state["door_options"]), create_input_field("Count (Per Elevation)", "door_count")]),
            create_dropdown("Style", "door_stile", state["stile_options"]),
            ft.Text("Hardware:", size=12, color=COLOR_TEXT_DIM),
            ft.Column([cb for cb in hardware_cbs.values()], spacing=0),
            ft.Row([
                ft.ElevatedButton("ADD DOOR", bgcolor="#4CAF50", color="white", on_click=lambda e: modify_door("add"), expand=True),
                ft.ElevatedButton("UPDATE DOOR", bgcolor="#FF9800", color="white", on_click=lambda e: modify_door("update"), expand=True),
            ]),
            door_list_col
        ], spacing=2)
        
        door_manager_expanded = False  # Default to collapsed
        door_manager_content_container = ft.Container(
            content=door_manager_content,
            visible=False,  # Hidden by default
            animate=ft.Animation(300, "easeOut"),
            padding=0
        )
        
        def toggle_door_manager(e):
            nonlocal door_manager_expanded
            door_manager_expanded = not door_manager_expanded
            door_manager_content_container.visible = door_manager_expanded
            door_manager_icon.icon = ft.Icons.EXPAND_MORE if door_manager_expanded else ft.Icons.CHEVRON_RIGHT
            page.update()
        
        door_manager_icon = ft.IconButton(
            icon=ft.Icons.CHEVRON_RIGHT,  # Pointing right when collapsed
            icon_color=COLOR_ACCENT,
            on_click=toggle_door_manager,
            tooltip="Toggle Door Manager"
        )
        
        door_col = ft.Column([
            ft.Row([
                ft.Text("DOOR MANAGER", size=14, weight="bold", color=COLOR_ACCENT, expand=True),
                door_manager_icon
            ]),
            door_manager_content_container
        ], spacing=0)
        
        form_col = ft.Column([
            ft.Text("ELEVATION DETAILS", size=14, weight="bold", color=COLOR_ACCENT),
            # Door only at top - when checked, hides rest of form (still requires elevation name)
            assign_ref("door_only_checkbox", ft.Checkbox(
                label="Door only (no bays) — still requires elevation name",
                value=False,
                fill_color=COLOR_ACCENT,
                on_change=update_door_only_visibility
            )),
            ft.Row([create_input_field("Elevation Type (Name)", "type")]),

            # Rest of form - hidden when door_only
            assign_ref("door_only_rest_container", ft.Column([
                ft.Row([create_dropdown("System", "system", state["system_options"], on_change=update_bay_visibility), 
                       create_dropdown("Finish", "finish", state["finish_options"])]),
                ft.Row([create_input_field("Quantity", "count", on_change=update_qty_per_elev_visibility)]),
                
                # Quantity Per Elevation Toggle (shown only when count > 1)
                assign_ref("qty_per_elev_container", ft.Container(
                    content=ft.Row([qty_per_elev_checkbox]),
                    visible=False,
                    margin=ft.margin.only(top=5, bottom=5)
                )),
                
                # Cost Per Elevation Toggles (shown only when count > 1)
                assign_ref("cost_per_elev_container", ft.Container(
                    content=ft.Column([
                        total_cost_per_elev_checkbox,
                        discounted_cost_per_elev_checkbox
                    ], spacing=5),
                    visible=False,
                    margin=ft.margin.only(top=5, bottom=5)
                )),
                
                # Dimensions
                ft.Container(content=ft.Column([
                    ft.Text("DIMENSIONS", size=12, weight="bold", color=COLOR_TEXT_DIM),
                    ft.Row([
                        create_input_field("Opening Width (\")", "width", on_change=lambda e: auto_refresh_bay_diagram()), 
                        create_input_field("Opening Height (\")", "height", on_change=lambda e: auto_refresh_bay_diagram())
                    ]),
                ]), margin=ft.margin.only(top=10)),

                # Bays (Hidden by default if not Yes45 or door-only)
                assign_ref("bay_config_container", ft.Container(content=ft.Column([
                    ft.Text("BAY CONFIGURATION", size=12, weight="bold", color=COLOR_TEXT_DIM),
                    ft.Row([
                        create_input_field("Bays Wide", "bays_wide", numeric=True, on_change=lambda e: (update_dynamic_bay_inputs(e), auto_refresh_bay_diagram())), 
                        create_input_field("Bays Tall", "bays_tall", numeric=True, on_change=lambda e: (update_dynamic_bay_inputs(e), auto_refresh_bay_diagram()))
                    ]),
                ]), margin=ft.margin.only(top=10))),
                
                # Containers for dynamic inputs
                assign_ref("custom_w_container", ft.Container(
                    content=inputs.setdefault("custom_w_col", ft.Column([], spacing=10)),
                    visible=False
                )),
                assign_ref("custom_h_container", ft.Container(
                    content=inputs.setdefault("custom_h_col", ft.Column([], spacing=10)),
                    visible=False
                )),
            ], spacing=0)),

            # Door Manager (always visible)
            door_col,
            
            # Create/Update/Delete buttons (below door manager, very close)
            ft.Container(
                content=ft.Row([
                    assign_ref("save_btn", ft.ElevatedButton("CREATE ELEVATION", bgcolor=COLOR_ACCENT, color="white", on_click=save_elevation_action, expand=True, height=50)),
                    assign_ref("duplicate_btn", ft.IconButton(ft.Icons.CONTENT_COPY, icon_color=COLOR_ACCENT, tooltip="Duplicate Elevation", on_click=duplicate_elevation_action, visible=False)),
                    ft.IconButton(ft.Icons.DELETE_FOREVER, icon_color="red", tooltip="Delete Elevation", on_click=delete_elevation_action)
                ]),
                margin=ft.margin.only(top=2)
            )
        ], scroll=ft.ScrollMode.AUTO, expand=True)

        # Right Col: Bay Diagram (Full Right Half)
        bay_diagram_image = inputs.setdefault("bay_diagram_image", ft.Image(
            src_base64="",
            fit=ft.ImageFit.CONTAIN,
            visible=False
        ))
        
        assign_ref("bay_diagram_container", ft.Container(
            content=ft.Column([
                ft.Text("BAY DIAGRAM", size=16, weight="bold", color=COLOR_ACCENT, text_align=ft.TextAlign.CENTER),
                ft.Container(
                    content=bay_diagram_image,
                    expand=True,
                    alignment=ft.alignment.center,
                    padding=10
                )
            ], 
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=10,
            expand=True),
            visible=True,  # Container always visible, image visibility controls display
            expand=True,
            padding=20,
            bgcolor=COLOR_SURFACE,
            border_radius=10,
            border=ft.Border(
                ft.BorderSide(2, COLOR_ACCENT_LIGHT),
                ft.BorderSide(2, COLOR_ACCENT_LIGHT),
                ft.BorderSide(2, COLOR_ACCENT_LIGHT),
                ft.BorderSide(2, COLOR_ACCENT_LIGHT)
            )
        ))

        # Create Tabs
        def on_tab_change(e):
            """Handle tab change - refresh waste calculator when Summary tab is selected"""
            if e.control.selected_index == 1:  # Summary tab is index 1
                try:
                    refresh_waste_calculator()
                    page.update()
                except Exception as ex:
                    print(f"[WARNING] Error refreshing waste calculator on tab change: {ex}")
        
        tabs = ft.Tabs(
            selected_index=0,
            on_change=on_tab_change,
            tabs=[
                ft.Tab(
                    text="Elevations",
                    content=ft.Row([
                        ft.Container(
                            content=ft.Column([
                                saved_dd, 
                                ft.Divider(color="transparent"), 
                                form_col
                            ]),
                            expand=1, 
                            bgcolor=COLOR_SURFACE, 
                            border_radius=10, 
                            padding=20, 
                            margin=ft.margin.only(left=10, top=10, bottom=10)
                        ),
                        ft.Container(
                            content=inputs["bay_diagram_container"],
                            expand=1,
                            margin=ft.margin.only(right=10, top=10, bottom=10)
                        )
                    ], expand=True)
                ),
                ft.Tab(
                    text="Summary",
                    content=ft.Container(
                        content=summary_settings_container,
                        padding=20,
                        expand=True
                    )
                )
            ],
            expand=True
        )
        
        # Main Layout
        main_content = tabs

        # Initial visibility check
        # The view is returned but not yet added to page.views in route_change,
        # so page.update() in update_bay_visibility fails if called directly.
        # Instead, we just set the initial visibility state of the controls.
        
        is_yes45_init = inputs["system"].value == "YES 45TU FRONT SET(OG)"
        door_only_init = inputs.get("door_only_checkbox") and inputs["door_only_checkbox"].value
        show_bays_init = is_yes45_init and not door_only_init
        
        # Safely set initial visibility on containers if they exist
        if inputs.get("bay_config_container"):
            inputs["bay_config_container"].visible = show_bays_init
        if inputs.get("bays_wide") and inputs["bays_wide"].parent:
             inputs["bays_wide"].parent.visible = show_bays_init
        if inputs.get("bays_tall") and inputs["bays_tall"].parent:
             inputs["bays_tall"].parent.visible = show_bays_init
             
        if inputs.get("custom_w_container"): inputs["custom_w_container"].visible = show_bays_init
        if inputs.get("custom_h_container"): inputs["custom_h_container"].visible = show_bays_init
        # Bay diagram visibility will be handled by auto_refresh_bay_diagram
        
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
        # Smooth transition - only rebuild what's needed
        target_route = page.route
        
        if target_route == "/":
            page.views.clear()
            page.views.append(build_projects_view())
        elif target_route == "/workspace" and state["current_project"]:
            # Keep projects view as base, add workspace on top
            if not page.views or page.views[0].route != "/":
                page.views.clear()
                page.views.append(build_projects_view())
            # Remove any views after the first one
            while len(page.views) > 1:
                page.views.pop()
            page.views.append(build_workspace_view())
        elif target_route == "/ml_analytics":
            if not page.views or page.views[0].route != "/":
                page.views.clear()
                page.views.append(build_projects_view())
            while len(page.views) > 1:
                page.views.pop()
            page.views.append(build_ml_analytics_view())
        else:
            # Default fallback
            page.views.clear()
            page.views.append(build_projects_view())
        
        page.update()

    def view_pop(e):
        if len(page.views) > 1:
            page.views.pop()
            top_view = page.views[-1]
            page.go(top_view.route)
        else:
            # If no views left, go to home
            page.go("/")

    page.on_route_change = route_change
    page.on_view_pop = view_pop
    page.go(page.route)

if __name__ == "__main__":
    # Get the correct assets directory path (works for both dev and PyInstaller)
    assets_dir_path = get_assets_dir()
    ft.app(target=main, assets_dir=assets_dir_path)
