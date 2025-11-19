"""Path management utilities for project file paths."""
import os
from openpyxl import Workbook

PROJECTS_DIR = ".files"
MASTER_PROJECT_LIST_FILE = os.path.join(PROJECTS_DIR, "projects_list.json")


def sanitize_name(name):
    """Sanitize a name for use in file paths."""
    if not name:
        return name
    return name.replace(" ", "_").replace("/", "_").replace("\\", "_")


def get_project_base_path(project_name, projects_dir=PROJECTS_DIR):
    """Get the base path for a project."""
    if not project_name:
        return None
    sanitized_name = sanitize_name(project_name)
    return os.path.join(projects_dir, sanitized_name)


def get_excel_path(project_name, projects_dir=PROJECTS_DIR):
    """Get the Excel report path for a project."""
    if not project_name:
        return os.path.join(projects_dir, "default_report.xlsx")
    base_path = get_project_base_path(project_name, projects_dir)
    return f"{base_path}_Report.xlsx"


def get_elevations_json_path(project_name, projects_dir=PROJECTS_DIR):
    """Get the elevations JSON path for a project."""
    if not project_name:
        return os.path.join(projects_dir, "default_elevations.json")
    base_path = get_project_base_path(project_name, projects_dir)
    return f"{base_path}_Elevations.json"


def get_extra_materials_json_path(project_name, projects_dir=PROJECTS_DIR):
    """Get the extra materials JSON path for a project."""
    if not project_name:
        return os.path.join(projects_dir, "default_extra_materials.json")
    base_path = get_project_base_path(project_name, projects_dir)
    return f"{base_path}_ExtraMaterials.json"


def get_door_json_path(project_name, elev_type, projects_dir=PROJECTS_DIR):
    """Get the door JSON path for a specific elevation."""
    if not project_name or not elev_type:
        return None
    base_path = get_project_base_path(project_name, projects_dir)
    safe_elev_type = sanitize_name(elev_type)
    return f"{base_path}_{safe_elev_type}_doors.json"


def get_project_paths(project_name, projects_dir=PROJECTS_DIR):
    """Get all project-related file paths as a dictionary."""
    return {
        'excel_path': get_excel_path(project_name, projects_dir),
        'elevations_json_path': get_elevations_json_path(project_name, projects_dir),
        'extra_materials_json_path': get_extra_materials_json_path(project_name, projects_dir),
    }


def ensure_excel_file_exists(excel_path):
    """Ensure the Excel file exists, creating it if it doesn't."""
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Report"
        wb.save(excel_path)


def get_unique_report_path(project_name, reports_dir=None):
    """Generate a unique report path with timestamp."""
    import datetime
    
    if reports_dir is None:
        # Get the directory of the calling script (main.py)
        import inspect
        frame = inspect.currentframe()
        try:
            caller_file = frame.f_back.f_globals.get('__file__', '')
            reports_dir = os.path.join(os.path.dirname(os.path.abspath(caller_file)), 'reports')
        finally:
            del frame
    
    os.makedirs(reports_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_project_name = sanitize_name(project_name)
    return os.path.join(reports_dir, f"{sanitized_project_name}_Report_{timestamp}.xlsx")


def get_project_files_to_delete(project_name, projects_dir=PROJECTS_DIR):
    """Get list of all files that should be deleted for a project."""
    base_path = get_project_base_path(project_name, projects_dir)
    if not base_path:
        return []
    
    files_to_delete = [
        f"{base_path}_Report.xlsx",
        f"{base_path}_Elevations.json",
        f"{base_path}_ExtraMaterials.json"
    ]
    
    # Find all door files for this project
    base_name = os.path.basename(base_path)
    if os.path.exists(projects_dir):
        for file_path in os.listdir(projects_dir):
            if file_path.startswith(base_name) and "_doors.json" in file_path:
                files_to_delete.append(os.path.join(projects_dir, file_path))
    
    return files_to_delete

