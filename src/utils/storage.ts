// Storage utilities for managing projects, elevations, and extra materials
// Uses localStorage for browser-based storage

const PROJECTS_KEY = 'ug_projects_list';
const PROJECT_PREFIX = 'ug_project_';
const ELEVATIONS_SUFFIX = '_Elevations';
const MATERIALS_SUFFIX = '_ExtraMaterials';
const DOORS_SUFFIX = '_doors';

export interface ProjectData {
  name: string;
  elevations: Record<string, ElevationData>;
  extraMaterials: Record<string, any>;
}

export interface ElevationData {
  system: string;
  finish: string;
  total_count: number;
  opening_width_inches: number;
  opening_height_inches: number;
  sqft_per_type: number;
  total_sqft: number;
  perimeter_ft: number;
  total_perimeter_ft: number;
  bays_wide?: number;
  bays_tall?: number;
  custom_bay_widths?: number[];
  custom_bay_heights?: number[];
  calculated_outputs: any[];
  material_impact?: any[];
}

export function loadProjects(): string[] {
  try {
    const data = localStorage.getItem(PROJECTS_KEY);
    return data ? JSON.parse(data) : [];
  } catch {
    return [];
  }
}

export function saveProjects(projects: string[]): void {
  try {
    localStorage.setItem(PROJECTS_KEY, JSON.stringify(projects));
  } catch (e) {
    console.error('Error saving projects:', e);
  }
}

export function loadElevations(projectName: string): Record<string, ElevationData> {
  try {
    const key = `${PROJECT_PREFIX}${projectName}${ELEVATIONS_SUFFIX}`;
    const data = localStorage.getItem(key);
    return data ? JSON.parse(data) : {};
  } catch {
    return {};
  }
}

export function saveElevations(projectName: string, elevations: Record<string, ElevationData>): void {
  try {
    const key = `${PROJECT_PREFIX}${projectName}${ELEVATIONS_SUFFIX}`;
    localStorage.setItem(key, JSON.stringify(elevations));
  } catch (e) {
    console.error('Error saving elevations:', e);
  }
}

export function loadExtraMaterials(projectName: string): Record<string, any> {
  try {
    const key = `${PROJECT_PREFIX}${projectName}${MATERIALS_SUFFIX}`;
    const data = localStorage.getItem(key);
    return data ? JSON.parse(data) : {};
  } catch {
    return {};
  }
}

export function saveExtraMaterials(projectName: string, materials: Record<string, any>): void {
  try {
    const key = `${PROJECT_PREFIX}${projectName}${MATERIALS_SUFFIX}`;
    localStorage.setItem(key, JSON.stringify(materials));
  } catch (e) {
    console.error('Error saving extra materials:', e);
  }
}

export function loadDoors(projectName: string, elevType: string): any[] {
  try {
    const key = `${PROJECT_PREFIX}${projectName}_${elevType}${DOORS_SUFFIX}`;
    const data = localStorage.getItem(key);
    return data ? JSON.parse(data) : [];
  } catch {
    return [];
  }
}

export function saveDoors(projectName: string, elevType: string, doors: any[]): void {
  try {
    const key = `${PROJECT_PREFIX}${projectName}_${elevType}${DOORS_SUFFIX}`;
    localStorage.setItem(key, JSON.stringify(doors));
  } catch (e) {
    console.error('Error saving doors:', e);
  }
}

export function deleteProject(projectName: string): void {
  try {
    // Remove from projects list
    const projects = loadProjects();
    const updated = projects.filter(p => p !== projectName);
    saveProjects(updated);

    // Remove all project-related data
    const keys = [
      `${PROJECT_PREFIX}${projectName}${ELEVATIONS_SUFFIX}`,
      `${PROJECT_PREFIX}${projectName}${MATERIALS_SUFFIX}`,
    ];

    // Remove door files (need to check all keys)
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith(`${PROJECT_PREFIX}${projectName}_`) && key.endsWith(DOORS_SUFFIX)) {
        localStorage.removeItem(key);
      }
    }

    keys.forEach(key => localStorage.removeItem(key));
  } catch (e) {
    console.error('Error deleting project:', e);
  }
}

