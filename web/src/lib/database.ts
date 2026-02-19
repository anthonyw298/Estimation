import { createClient, SupabaseClient } from '@supabase/supabase-js';
import { ElevationData, DoorConfig, ProjectSettings, ExtraMaterial } from '@/types';

const SUPABASE_URL = process.env.NEXT_PUBLIC_SUPABASE_URL!;
const SUPABASE_KEY = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!;

if (!SUPABASE_URL || !SUPABASE_KEY) {
  throw new Error('Missing NEXT_PUBLIC_SUPABASE_URL or NEXT_PUBLIC_SUPABASE_ANON_KEY environment variables');
}

const supabase: SupabaseClient = createClient(SUPABASE_URL, SUPABASE_KEY);
export { supabase };

export const db = {
  // ==================== PROJECTS ====================

  async getProjects(): Promise<string[]> {
    const { data, error } = await supabase
      .from('projects')
      .select('name');
    if (error) throw error;
    return (data ?? []).map((p: { name: string }) => p.name);
  },

  async deleteProject(name: string): Promise<void> {
    // Delete project and all related data (same order as Python)
    const { error: e1 } = await supabase.from('projects').delete().eq('name', name);
    if (e1) throw e1;
    const { error: e2 } = await supabase.from('elevations').delete().eq('project_name', name);
    if (e2) throw e2;
    const { error: e3 } = await supabase.from('settings').delete().eq('project_name', name);
    if (e3) throw e3;
    const { error: e4 } = await supabase.from('doors').delete().eq('project_name', name);
    if (e4) throw e4;
    const { error: e5 } = await supabase.from('materials').delete().eq('project_name', name);
    if (e5) throw e5;
  },

  async createProject(name: string): Promise<void> {
    const { error } = await supabase
      .from('projects')
      .insert({ name });
    if (error) throw error;
  },

  // ==================== ELEVATIONS ====================

  async getElevations(projectName: string): Promise<Record<string, ElevationData>> {
    const { data, error } = await supabase
      .from('elevations')
      .select('*')
      .eq('project_name', projectName);
    if (error) throw error;

    const elevations: Record<string, ElevationData> = {};
    for (const row of data ?? []) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const raw = row.data as any;

      // Normalize Python key names → TypeScript key names
      if (raw && raw.system !== undefined && raw.system_type === undefined) {
        raw.system_type = raw.system;
      }
      if (raw && raw.material_impact !== undefined && raw.material_impacts === undefined) {
        raw.material_impacts = raw.material_impact;
      }

      elevations[row.name] = raw as ElevationData;
    }
    return elevations;
  },

  async saveElevations(projectName: string, elevations: Record<string, ElevationData>): Promise<void> {
    for (const [name, data] of Object.entries(elevations)) {
      const { error } = await supabase
        .from('elevations')
        .upsert(
          { project_name: projectName, name, data },
          { onConflict: 'project_name,name' }
        );
      if (error) throw error;
    }
  },

  async saveElevation(projectName: string, elevationName: string, data: ElevationData): Promise<void> {
    const { error } = await supabase
      .from('elevations')
      .upsert(
        { project_name: projectName, name: elevationName, data },
        { onConflict: 'project_name,name' }
      );
    if (error) throw error;
  },

  async deleteElevation(projectName: string, elevationName: string): Promise<void> {
    const { error: e1 } = await supabase
      .from('elevations')
      .delete()
      .eq('project_name', projectName)
      .eq('name', elevationName);
    if (e1) throw e1;

    // Also delete associated doors (same as Python)
    const { error: e2 } = await supabase
      .from('doors')
      .delete()
      .eq('project_name', projectName)
      .eq('elevation_name', elevationName);
    if (e2) throw e2;
  },

  // ==================== SETTINGS ====================

  async getSettings(projectName: string): Promise<ProjectSettings> {
    const { data, error } = await supabase
      .from('settings')
      .select('data')
      .eq('project_name', projectName);
    if (error) throw error;

    if (data && data.length > 0) {
      return data[0].data as ProjectSettings;
    }
    return {} as ProjectSettings;
  },

  async saveSettings(projectName: string, settings: ProjectSettings): Promise<void> {
    const { error } = await supabase
      .from('settings')
      .upsert(
        { project_name: projectName, data: settings },
        { onConflict: 'project_name' }
      );
    if (error) throw error;
  },

  // ==================== DOORS ====================

  async getAllDoors(projectName: string): Promise<Record<string, DoorConfig[]>> {
    const { data, error } = await supabase
      .from('doors')
      .select('*')
      .eq('project_name', projectName);
    if (error) throw error;
    const result: Record<string, DoorConfig[]> = {};
    for (const row of data ?? []) {
      result[row.elevation_name] = row.data as DoorConfig[];
    }
    return result;
  },

  async getDoors(projectName: string, elevationName: string): Promise<DoorConfig[]> {
    const { data, error } = await supabase
      .from('doors')
      .select('data')
      .eq('project_name', projectName)
      .eq('elevation_name', elevationName);
    if (error) throw error;

    if (data && data.length > 0) {
      return data[0].data as DoorConfig[];
    }
    return [];
  },

  async saveDoors(projectName: string, elevationName: string, doors: DoorConfig[]): Promise<void> {
    const { error } = await supabase
      .from('doors')
      .upsert(
        { project_name: projectName, elevation_name: elevationName, data: doors },
        { onConflict: 'project_name,elevation_name' }
      );
    if (error) throw error;
  },

  // ==================== MATERIALS ====================

  async getMaterials(projectName: string): Promise<Record<string, ExtraMaterial>> {
    const { data, error } = await supabase
      .from('materials')
      .select('data')
      .eq('project_name', projectName);
    if (error) throw error;

    if (data && data.length > 0) {
      return data[0].data as Record<string, ExtraMaterial>;
    }
    return {};
  },

  async saveMaterials(projectName: string, materials: Record<string, ExtraMaterial>): Promise<void> {
    const { error } = await supabase
      .from('materials')
      .upsert(
        { project_name: projectName, data: materials },
        { onConflict: 'project_name' }
      );
    if (error) throw error;
  },

  // ==================== ML TRAINING DATA ====================
  // Stored in settings table with special key '__ml_training_data__'

  async getMLData(): Promise<unknown[]> {
    const { data, error } = await supabase
      .from('settings')
      .select('data')
      .eq('project_name', '__ml_training_data__');
    if (error) throw error;

    if (data && data.length > 0) {
      return (data[0].data as { samples: unknown[] })?.samples ?? [];
    }
    return [];
  },

  async saveMLData(samples: unknown[]): Promise<void> {
    const { error } = await supabase
      .from('settings')
      .upsert(
        { project_name: '__ml_training_data__', data: { samples } },
        { onConflict: 'project_name' }
      );
    if (error) throw error;
  },
};
