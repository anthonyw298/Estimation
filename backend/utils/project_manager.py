"""Project management operations."""
import os
from utils.file_operations import (
    load_project_list,
    save_project_list,
    ensure_json_file_exists,
    create_excel_file
)
from utils.path_manager import (
    get_project_paths,
    get_project_files_to_delete,
    PROJECTS_DIR
)


def initialize_project_files(project_name):
    """Initialize all required files for a project."""
    paths = get_project_paths(project_name)
    
    # Ensure Excel file exists
    create_excel_file(paths['excel_path'])
    
    # Ensure extra materials JSON exists
    ensure_json_file_exists(paths['extra_materials_json_path'], default_data={})
    
    return paths


def create_project(project_name, existing_projects):
    """Create a new project.
    
    Returns:
        tuple: (success: bool, error_message: str or None, updated_projects_list: list)
    """
    new_name = project_name.strip()
    
    if not new_name:
        return False, "Please enter a name for the new project.", existing_projects
    
    if new_name in existing_projects:
        return False, f"Project '{new_name}' already exists.", existing_projects
    
    # Add to project list
    updated_projects = existing_projects + [new_name]
    save_project_list(updated_projects)
    
    # Initialize project files
    initialize_project_files(new_name)
    
    return True, None, updated_projects


def delete_project(project_name, existing_projects):
    """Delete a project and all its associated files.
    
    Returns:
        tuple: (success: bool, error_message: str or None, updated_projects_list: list, next_project: str or None)
    """
    if not project_name:
        return False, "No project selected to delete.", existing_projects, None
    
    if project_name not in existing_projects:
        return False, f"Project '{project_name}' not found.", existing_projects, None
    
    try:
        # Get all files to delete
        files_to_delete = get_project_files_to_delete(project_name)
        
        # Delete all files
        for file_path in files_to_delete:
            if os.path.exists(file_path):
                os.remove(file_path)
        
        # Remove from project list
        updated_projects = [p for p in existing_projects if p != project_name]
        save_project_list(updated_projects)
        
        # Determine next project to select
        next_project = updated_projects[0] if updated_projects else None
        
        return True, None, updated_projects, next_project
        
    except Exception as e:
        return False, f"Error deleting project '{project_name}': {e}", existing_projects, None


def load_projects():
    """Load the list of projects.
    
    Returns:
        list: List of project names
    """
    try:
        return load_project_list()
    except Exception:
        return []

