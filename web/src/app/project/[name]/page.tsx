'use client';

import { useState, useEffect, useCallback, use } from 'react';
import { useRouter } from 'next/navigation';
import { db } from '@/lib/database';
import { ElevationData, DoorConfig, ProjectSettings, ExtraMaterial } from '@/types';
import { reverseMaterialImpact } from '@/lib/pricing';
import ElevationEditor from '@/components/ElevationEditor';
import WasteAnalysis from '@/components/WasteAnalysis';
import PricingAdjustmentTab from '@/components/PricingAdjustmentTab';
import ReportOptionsDialog from '@/components/ReportOptionsDialog';
import {
  ArrowLeft,
  Plus,
  Trash2,
  Copy,
  ChevronRight,
  Layers,
  FileSpreadsheet,
  BarChart3,
  Loader2,
  X,
  Settings,
  Settings2,
} from 'lucide-react';

interface ProjectPageProps {
  params: Promise<{ name: string }>;
}

export default function ProjectPage({ params }: ProjectPageProps) {
  const { name } = use(params);
  const projectName = decodeURIComponent(name);
  const router = useRouter();

  // Data state
  const [elevations, setElevations] = useState<Record<string, ElevationData>>({});
  const [doors, setDoors] = useState<Record<string, DoorConfig[]>>({});
  const [settings, setSettings] = useState<ProjectSettings>({});
  const [materials, setMaterials] = useState<Record<string, ExtraMaterial>>({});

  // UI state
  const [selectedElevation, setSelectedElevation] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [showAddElevation, setShowAddElevation] = useState(false);
  const [newElevationName, setNewElevationName] = useState('');
  const [confirmDelete, setConfirmDelete] = useState<string | null>(null);
  const [duplicateElevation, setDuplicateElevation] = useState<string | null>(null);
  const [duplicateName, setDuplicateName] = useState('');
  const [showWasteAnalysis, setShowWasteAnalysis] = useState(false);
  const [activeTab, setActiveTab] = useState<'editor' | 'waste'>('editor');
  const [workspaceTab, setWorkspaceTab] = useState<'elevations' | 'pricing'>('elevations');
  const [showReportOptions, setShowReportOptions] = useState(false);

  // Load project data
  useEffect(() => {
    loadProjectData();
  }, [projectName]);

  async function loadProjectData() {
    try {
      setLoading(true);
      const [elevationData, settingsData, materialsData, doorsData] = await Promise.all([
        db.getElevations(projectName),
        db.getSettings(projectName),
        db.getMaterials(projectName),
        db.getAllDoors(projectName),
      ]);
      setElevations(elevationData || {});
      setSettings(settingsData || {});
      setMaterials(materialsData || {});
      setDoors(doorsData || {});
    } catch (error) {
      console.error('Failed to load project data:', error);
    } finally {
      setLoading(false);
    }
  }

  const handleElevationUpdate = useCallback(
    async (elevationName: string, data: ElevationData) => {
      setElevations((prev) => ({ ...prev, [elevationName]: data }));
      try {
        setSaving(true);
        await db.saveElevation(projectName, elevationName, data);
      } catch (error) {
        console.error('Failed to save elevation:', error);
      } finally {
        setSaving(false);
      }
    },
    [projectName]
  );

  const handleDoorsUpdate = useCallback(
    async (elevationName: string, doorConfigs: DoorConfig[]) => {
      setDoors((prev) => ({ ...prev, [elevationName]: doorConfigs }));
      try {
        await db.saveDoors(projectName, elevationName, doorConfigs);
      } catch (error) {
        console.error('Failed to save doors:', error);
      }
    },
    [projectName]
  );

  const handleSettingsUpdate = useCallback(
    async (newSettings: ProjectSettings) => {
      setSettings(newSettings);
      try {
        await db.saveSettings(projectName, newSettings);
      } catch (error) {
        console.error('Failed to save settings:', error);
      }
    },
    [projectName]
  );

  const handleMaterialsUpdate = useCallback(
    async (newMaterials: Record<string, ExtraMaterial>) => {
      setMaterials(newMaterials);
      try {
        await db.saveMaterials(projectName, newMaterials);
      } catch (error) {
        console.error('Failed to save materials:', error);
      }
    },
    [projectName]
  );

  const handleResetInventory = useCallback(async () => {
    // 1. Clear all materials
    const emptyMaterials: Record<string, ExtraMaterial> = {};
    setMaterials(emptyMaterials);

    // 2. Clear only material_impacts from elevations (keep calculated_outputs
    //    and single_elevation_outputs so prices/exports still work and the
    //    sidebar doesn't show "needs calculation").
    const cleanedElevations: Record<string, ElevationData> = {};
    for (const [elevName, elev] of Object.entries(elevations)) {
      const { material_impacts, ...rest } = elev;
      cleanedElevations[elevName] = rest as ElevationData;
    }
    setElevations(cleanedElevations);

    // 3. Persist to database
    try {
      setSaving(true);
      await db.saveMaterials(projectName, emptyMaterials);
      await Promise.all(
        Object.entries(cleanedElevations).map(([eName, eData]) =>
          db.saveElevation(projectName, eName, eData)
        )
      );
    } catch (error) {
      console.error('Failed to reset inventory:', error);
    } finally {
      setSaving(false);
    }
  }, [projectName, elevations]);

  async function handleAddElevation() {
    const trimmed = newElevationName.trim();
    if (!trimmed) return;

    if (elevations[trimmed]) {
      alert('An elevation with this name already exists.');
      return;
    }

    const newElevation: ElevationData = {
      system_type: 'YES 45TU Front Set (OG)',
      finish: 'Clear',
      opening_width_inches: 0,
      opening_height_inches: 0,
      bays_wide: 0,
      bays_tall: 0,
      total_count: 0,
    };

    setElevations((prev) => ({ ...prev, [trimmed]: newElevation }));
    setSelectedElevation(trimmed);
    setNewElevationName('');
    setShowAddElevation(false);

    try {
      await db.saveElevation(projectName, trimmed, newElevation);
    } catch (error) {
      console.error('Failed to save new elevation:', error);
    }
  }

  async function handleDeleteElevation(elevationName: string) {
    // CRITICAL: Reverse material impacts from the deleted elevation so
    // leftovers it generated are removed from the shared materials inventory
    const elevToDelete = elevations[elevationName];
    if (elevToDelete?.material_impacts && elevToDelete.material_impacts.length > 0) {
      const materialsClone: Record<string, ExtraMaterial> = {};
      for (const [k, v] of Object.entries(materials)) {
        materialsClone[k] = {
          quantity: v.quantity,
          length_pieces: [...v.length_pieces],
        };
      }
      reverseMaterialImpact(elevToDelete.material_impacts, materialsClone);
      // Persist the reversed materials
      setMaterials(materialsClone);
      try {
        await db.saveMaterials(projectName, materialsClone);
      } catch (_) {
        // best-effort
      }
    }

    const updated = { ...elevations };
    delete updated[elevationName];
    setElevations(updated);

    const updatedDoors = { ...doors };
    delete updatedDoors[elevationName];
    setDoors(updatedDoors);

    if (selectedElevation === elevationName) {
      const remaining = Object.keys(updated);
      setSelectedElevation(remaining.length > 0 ? remaining[0] : null);
    }
    setConfirmDelete(null);

    try {
      await db.deleteElevation(projectName, elevationName);
    } catch (error) {
      console.error('Failed to delete elevation:', error);
    }
  }

  async function handleDuplicateElevation() {
    if (!duplicateElevation) return;
    const newName = duplicateName.trim();
    if (!newName) return;

    if (elevations[newName]) {
      alert('An elevation with this name already exists.');
      return;
    }

    // Deep clone the elevation data, strip material impacts and calculated outputs for a fresh copy
    const original = elevations[duplicateElevation];
    const cloned: ElevationData = JSON.parse(JSON.stringify(original));
    delete cloned.material_impacts;
    delete cloned.calculated_outputs;
    delete cloned.single_elevation_outputs;

    // Clone door data if it exists
    const originalDoors = doors[duplicateElevation];
    const clonedDoors: DoorConfig[] = originalDoors
      ? JSON.parse(JSON.stringify(originalDoors))
      : [];

    // Update local state
    setElevations((prev) => ({ ...prev, [newName]: cloned }));
    if (clonedDoors.length > 0) {
      setDoors((prev) => ({ ...prev, [newName]: clonedDoors }));
    }
    setSelectedElevation(newName);
    setDuplicateElevation(null);
    setDuplicateName('');

    // Persist to database
    try {
      await db.saveElevation(projectName, newName, cloned);
      if (clonedDoors.length > 0) {
        await db.saveDoors(projectName, newName, clonedDoors);
      }
    } catch (error) {
      console.error('Failed to duplicate elevation:', error);
    }
  }

  async function handleExportExcel() {
    try {
      const { exportToExcel } = await import('@/lib/export');
      await exportToExcel(projectName, elevations, doors, settings, materials);
    } catch (error) {
      console.error('Failed to export:', error);
      alert('Export failed. Please try again.');
    }
  }

  // Compute per-elevation discounted cost (matching Excel per-elevation sheet).
  // Each elevation is treated individually — no residual/waste deducted.
  // The discount multiplier is determined from the project-wide list total.
  function getElevationCost(elevationName: string): { list: number; discounted: number } | null {
    const elev = elevations[elevationName];
    if (!elev?.calculated_outputs) return null;

    const DISCOUNTABLE = new Set(['profiles', 'accessories', 'gaskets']);
    const GASKETS = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

    function classify(item: { part_number: string; description: string; type: string }): string {
      const pn = item.part_number || '';
      const desc = (item.description || '').toLowerCase();
      const tp = (item.type || '').toLowerCase();
      if (pn === 'GLASS_AREA' || tp === 'glass') return 'glass';
      if (pn === 'JOINTS_FAB_LABOR' || tp === 'joints_fab_labor' || tp === 'fabrication' ||
          desc.includes('joints fabrication') || desc.includes('fabrication labor')) return 'fabrication';
      if (tp === 'door' || tp === 'doors') return 'doors';
      if (tp === 'calculations') return 'calculations';
      if (desc.includes('gasket') || GASKETS.has(pn)) return 'gaskets';
      if (tp === 'accessory' || tp === 'accessories') return 'accessories';
      return 'profiles';
    }

    // Project-wide list total to determine discount tier
    let projectTotal = 0;
    for (const e of Object.values(elevations)) {
      if (!e.calculated_outputs) continue;
      for (const item of e.calculated_outputs) {
        if (classify(item) === 'calculations') continue;
        projectTotal += item.price ?? 0;
      }
    }

    const threshold = settings.discount_threshold ?? 50000;
    const multiplier = settings.discount_multiplier
      ?? (projectTotal < threshold
        ? (settings.discount_multiplier_low ?? 0.614)
        : (settings.discount_multiplier_high ?? 0.572));

    let listTotal = 0;
    let discountedTotal = 0;
    for (const item of elev.calculated_outputs) {
      const cat = classify(item);
      if (cat === 'calculations') continue;
      const price = item.price ?? 0;
      listTotal += price;
      discountedTotal += DISCOUNTABLE.has(cat) ? price * multiplier : price;
    }

    return { list: listTotal, discounted: discountedTotal };
  }

  const elevationNames = Object.keys(elevations);

  // Loading state
  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="flex flex-col items-center animate-fade-up opacity-0">
          <Loader2 className="w-8 h-8 text-[#3b82f6] animate-spin mb-4" />
          <p className="text-[#b0b3c0] text-sm">Loading project...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="h-screen flex flex-col overflow-hidden">
      {/* Header */}
      <header className="glass-strong gradient-border-top border-b border-[#1e1e2a]/40 flex-shrink-0">
        <div className="px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <button
              onClick={() => router.push('/')}
              className="p-2 rounded-xl hover:bg-[#111118] text-[#b0b3c0] hover:text-[#ffffff] transition-colors duration-200"
              title="Back to projects"
            >
              <ArrowLeft className="w-5 h-5" />
            </button>
            <div className="w-px h-6 bg-[#1e1e2a]" />
            <div className="flex items-center gap-2.5">
              <div className="relative">
                <div className="absolute inset-0 rounded-lg bg-blue-500/15 blur-md" />
                <div className="relative w-7 h-7 rounded-lg bg-gradient-to-br from-[#3b82f6]/15 to-[#6366f1]/10 border border-[#3b82f6]/20 flex items-center justify-center">
                  <Layers className="w-3.5 h-3.5 text-[#3b82f6]" />
                </div>
              </div>
              <h1 className="text-base font-semibold gradient-text-static tracking-tight truncate max-w-[300px]">
                {projectName}
              </h1>
            </div>
            {saving && (
              <span className="text-xs text-[#7a7c90] flex items-center gap-1.5 ml-2 animate-fade-in">
                <Loader2 className="w-3 h-3 animate-spin" />
                Saving...
              </span>
            )}
          </div>

          <div className="flex items-center gap-2">
            {/* Workspace tabs */}
            <div className="flex items-center bg-[#08080e] border border-[#1e1e2a] rounded-xl p-0.5 mr-2">
              {([
                { key: 'elevations' as const, label: 'Elevations', icon: <Layers className="w-3.5 h-3.5" /> },
                { key: 'pricing' as const, label: 'Pricing', icon: <Settings2 className="w-3.5 h-3.5" /> },
              ]).map(({ key, label, icon }) => (
                <button
                  key={key}
                  onClick={() => setWorkspaceTab(key)}
                  className={`flex items-center gap-1.5 px-4 py-2 text-xs font-semibold rounded-lg transition-colors duration-200 ${
                    workspaceTab === key
                      ? 'bg-gradient-to-r from-[#3b82f6] to-[#6366f1] text-white'
                      : 'text-[#7a7c90] hover:text-[#b0b3c0] hover:bg-[#111118]/50'
                  }`}
                >
                  {icon}
                  {label}
                </button>
              ))}
            </div>

            <div className="w-px h-5 bg-[#1e1e2a]" />

            <button
              onClick={() => setActiveTab(activeTab === 'waste' ? 'editor' : 'waste')}
              className={`flex items-center gap-2 px-3.5 py-2 text-sm font-semibold rounded-xl transition-colors duration-200 ${
                activeTab === 'waste'
                  ? 'bg-gradient-to-r from-[#3b82f6] to-[#6366f1] text-white'
                  : 'text-[#b0b3c0] hover:text-[#ffffff] hover:bg-[#111118]'
              }`}
            >
              <BarChart3 className="w-4 h-4" />
              Waste Analysis
            </button>
            <button
              onClick={() => setShowReportOptions(true)}
              className="flex items-center gap-2 px-3.5 py-2 text-sm font-medium text-[#b0b3c0] hover:text-[#ffffff] hover:bg-[#111118] rounded-xl transition-colors duration-200"
            >
              <FileSpreadsheet className="w-4 h-4" />
              Export
            </button>
          </div>
        </div>
      </header>

      {/* Body */}
      <div className="flex flex-1 overflow-hidden">
        {/* Sidebar - only visible on Elevations tab */}
        {workspaceTab === 'elevations' && (
          <aside className="w-[280px] flex-shrink-0 bg-[#08080e] border-r border-[#1e1e2a] flex flex-col overflow-hidden gradient-border-left">
            <div className="px-4 py-3 border-b border-[#1e1e2a] bg-[#0a0a12]">
              <h2 className="text-[10px] font-semibold text-[#7a7c90] uppercase tracking-[0.15em]">
                Elevations ({elevationNames.length})
              </h2>
            </div>

            {/* Elevation List */}
            <div className="flex-1 overflow-y-auto py-1.5">
              {elevationNames.length === 0 && (
                <div className="px-4 py-10 text-center animate-fade-up opacity-0">
                  <div className="w-12 h-12 rounded-xl bg-[#111118] border border-[#1e1e2a] flex items-center justify-center mx-auto mb-3">
                    <Layers className="w-6 h-6 text-[#2a2a3a]" />
                  </div>
                  <p className="text-xs text-[#7a7c90] leading-relaxed">
                    No elevations yet. Add one to get started.
                  </p>
                </div>
              )}

              {elevationNames.map((elevName) => {
                const isSelected = selectedElevation === elevName;
                const cost = getElevationCost(elevName);
                const isNew = !elevations[elevName]?.calculated_outputs || elevations[elevName].calculated_outputs!.length === 0;
                return (
                  <div
                    key={elevName}
                    className={`group relative flex items-center gap-2 px-3 py-2.5 mx-1.5 my-0.5 rounded-lg cursor-pointer transition-colors duration-200 ${
                      isSelected
                        ? 'bg-[#3b82f6]/10 sidebar-active'
                        : 'border-l-2 border-transparent hover:bg-[#0f0f17]'
                    }`}
                    onClick={() => {
                      setSelectedElevation(elevName);
                      setActiveTab('editor');
                    }}
                  >
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-1.5">
                        {isNew && (
                          <span className="badge bg-blue-500/15 text-blue-400">New</span>
                        )}
                        <p
                          className={`text-sm font-medium truncate transition-colors duration-200 ${
                            isSelected ? 'text-[#ffffff]' : 'text-[#b0b3c0]'
                          }`}
                        >
                          {elevName}
                        </p>
                      </div>
                      {isNew ? (
                        <p className="text-[10px] text-[#5a5c6e] mt-0.5 font-medium">
                          Click Update to calculate
                        </p>
                      ) : cost !== null ? (
                        <p className="text-xs text-[#7a7c90] mt-0.5 font-mono tabular-nums">
                          ${cost.discounted.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </p>
                      ) : null}
                    </div>

                    <div className="flex items-center gap-1">
                      {/* Duplicate button */}
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setDuplicateElevation(elevName);
                          setDuplicateName(`${elevName} (Copy)`);
                        }}
                        className="opacity-0 group-hover:opacity-100 p-1 rounded-md hover:bg-[#3b82f6]/10 text-[#5a5c6e] hover:text-[#3b82f6] transition-colors duration-200"
                        title="Duplicate elevation"
                      >
                        <Copy className="w-3.5 h-3.5" />
                      </button>
                      {/* Delete button */}
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setConfirmDelete(elevName);
                        }}
                        className="opacity-0 group-hover:opacity-100 p-1 rounded-md hover:bg-[#f87171]/10 text-[#5a5c6e] hover:text-[#f87171] transition-colors duration-200"
                        title="Delete elevation"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                      <ChevronRight
                        className={`w-3.5 h-3.5 transition-colors duration-200 ${
                          isSelected ? 'text-[#3b82f6]' : 'text-[#2a2a3a]'
                        }`}
                      />
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Add Elevation */}
            <div className="border-t border-[#1e1e2a] p-3 bg-[#0a0a12]/50">
              {showAddElevation ? (
                <div className="flex flex-col gap-2 animate-fade-in">
                  <input
                    type="text"
                    value={newElevationName}
                    onChange={(e) => setNewElevationName(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleAddElevation();
                      if (e.key === 'Escape') {
                        setShowAddElevation(false);
                        setNewElevationName('');
                      }
                    }}
                    placeholder="Elevation name"
                    className="w-full px-3 py-2 bg-[#0c0c12] border border-[#1e1e2a] rounded-lg text-sm text-[#ffffff] placeholder-[#5a5c6e] focus:outline-none focus:border-[#3b82f6] focus:ring-2 focus:ring-[#3b82f6]/20 transition-colors duration-200"
                    autoFocus
                  />
                  <div className="flex items-center gap-2">
                    <button
                      onClick={handleAddElevation}
                      disabled={!newElevationName.trim()}
                      className="flex-1 px-3 py-1.5 bg-gradient-to-r from-[#3b82f6] to-[#2563eb] hover:brightness-110 disabled:opacity-50 disabled:cursor-not-allowed text-white text-xs font-medium rounded-lg transition-colors duration-200"
                    >
                      Add
                    </button>
                    <button
                      onClick={() => {
                        setShowAddElevation(false);
                        setNewElevationName('');
                      }}
                      className="flex-1 px-3 py-1.5 text-[#b0b3c0] hover:text-[#ffffff] text-xs font-medium rounded-lg hover:bg-[#1e1e2a] transition-colors duration-200"
                    >
                      Cancel
                    </button>
                  </div>
                </div>
              ) : (
                <button
                  onClick={() => setShowAddElevation(true)}
                  className="w-full flex items-center justify-center gap-2 px-3 py-2.5 text-sm font-medium text-[#b0b3c0] hover:text-[#ffffff] hover:bg-[#3b82f6]/5 rounded-xl border border-dashed border-[#1e1e2a] hover:border-[#3b82f6]/40 transition-colors duration-200"
                >
                  <Plus className="w-4 h-4" />
                  Add Elevation
                </button>
              )}
            </div>
          </aside>
        )}

        {/* Main Content */}
        <main className="flex-1 overflow-y-auto">
          {/* Waste Analysis -- visible from ANY workspace tab when active */}
          {activeTab === 'waste' ? (
            <div className="p-6">
              <WasteAnalysis
                elevations={elevations}
                materials={materials}
                settings={settings}
                onResetInventory={handleResetInventory}
              />
            </div>
          ) : (
            <>
              {/* Elevations tab content */}
              {workspaceTab === 'elevations' && (
                <>
                  {selectedElevation && elevations[selectedElevation] ? (
                    <div className="p-6">
                      <ElevationEditor
                        key={selectedElevation}
                        projectName={projectName}
                        elevationName={selectedElevation}
                        elevationData={elevations[selectedElevation]}
                        doors={doors[selectedElevation] || []}
                        settings={settings}
                        materials={materials}
                        onSave={(name, data, doorConfigs, mats) => {
                          handleElevationUpdate(name, data);
                          handleDoorsUpdate(name, doorConfigs);
                          handleMaterialsUpdate(mats);
                        }}
                      />
                    </div>
                  ) : (
                    <div className="flex-1 flex items-center justify-center h-full min-h-[60vh]">
                      <div className="text-center animate-fade-up opacity-0">
                        <div className="w-16 h-16 rounded-2xl bg-[#111118] border border-[#1e1e2a] flex items-center justify-center mx-auto mb-6 shadow-lg shadow-black/20">
                          <Layers className="w-8 h-8 text-[#2a2a3a]" />
                        </div>
                        <h3 className="text-lg font-semibold text-[#ffffff] mb-2">
                          {elevationNames.length === 0
                            ? 'No elevations yet'
                            : 'Select an elevation'}
                        </h3>
                        <p className="text-sm text-[#7a7c90] max-w-xs mx-auto leading-relaxed">
                          {elevationNames.length === 0
                            ? 'Add your first elevation from the sidebar to start estimating.'
                            : 'Choose an elevation from the sidebar to view and edit its details.'}
                        </p>
                      </div>
                    </div>
                  )}
                </>
              )}

              {/* Pricing tab content (includes pricing config, additional costs, markups) */}
              {workspaceTab === 'pricing' && (
                <PricingAdjustmentTab
                  settings={settings}
                  onSettingsUpdate={handleSettingsUpdate}
                />
              )}
            </>
          )}
        </main>
      </div>

      {/* Delete Elevation Confirmation Modal */}
      {confirmDelete && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-md animate-overlay"
          onClick={() => setConfirmDelete(null)}
        >
          <div
            className="bg-[#111118] border border-[#1e1e2a] rounded-2xl w-full max-w-sm mx-4 p-7 shadow-2xl shadow-black/60 animate-scale-in"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-3 mb-4">
              <div className="w-11 h-11 rounded-full bg-[#f87171]/10 border border-[#f87171]/15 flex items-center justify-center flex-shrink-0">
                <Trash2 className="w-5 h-5 text-[#f87171]" />
              </div>
              <div>
                <h3 className="text-base font-semibold text-[#ffffff]">
                  Delete Elevation
                </h3>
                <p className="text-xs text-[#b0b3c0]">
                  This action cannot be undone
                </p>
              </div>
            </div>

            <p className="text-sm text-[#b0b3c0] mb-6 leading-relaxed">
              Are you sure you want to delete{' '}
              <span className="font-semibold text-[#ffffff]">{confirmDelete}</span>?
              All associated data will be permanently removed.
            </p>

            <div className="flex items-center gap-3 justify-end">
              <button
                onClick={() => setConfirmDelete(null)}
                className="px-4 py-2.5 text-sm font-medium text-[#b0b3c0] hover:text-[#ffffff] rounded-xl hover:bg-[#1e1e2a] transition-colors duration-200"
              >
                Cancel
              </button>
              <button
                onClick={() => handleDeleteElevation(confirmDelete)}
                className="flex items-center gap-2 px-4 py-2.5 bg-[#f87171] hover:bg-[#ef4444] text-white text-sm font-semibold rounded-xl transition-colors duration-200"
              >
                <Trash2 className="w-4 h-4" />
                Delete
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Duplicate Elevation Modal */}
      {duplicateElevation && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-md animate-overlay"
          onClick={() => { setDuplicateElevation(null); setDuplicateName(''); }}
        >
          <div
            className="bg-[#111118] border border-[#1e1e2a] rounded-2xl w-full max-w-sm mx-4 p-7 shadow-2xl shadow-black/60 animate-scale-in"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-3 mb-4">
              <div className="w-11 h-11 rounded-full bg-[#3b82f6]/10 border border-[#3b82f6]/15 flex items-center justify-center flex-shrink-0">
                <Copy className="w-5 h-5 text-[#3b82f6]" />
              </div>
              <div>
                <h3 className="text-base font-semibold text-[#ffffff]">
                  Duplicate Elevation
                </h3>
                <p className="text-xs text-[#8b8d9a]">
                  Copy of <span className="text-[#ffffff]">{duplicateElevation}</span>
                </p>
              </div>
            </div>

            <label className="block text-xs font-medium text-[#8b8d9a] mb-1.5">
              New elevation name
            </label>
            <input
              type="text"
              value={duplicateName}
              onChange={(e) => setDuplicateName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleDuplicateElevation();
                if (e.key === 'Escape') { setDuplicateElevation(null); setDuplicateName(''); }
              }}
              placeholder="Enter name..."
              className="w-full px-3.5 py-2.5 text-sm bg-[#0a0a12] border border-[#1e1e2a] rounded-xl text-[#eeeef2] placeholder-[#3e3f4d] focus:border-[#3b82f6]/50 focus:outline-none transition-colors duration-200 mb-5"
              autoFocus
            />

            <div className="flex items-center gap-3 justify-end">
              <button
                onClick={() => { setDuplicateElevation(null); setDuplicateName(''); }}
                className="px-4 py-2.5 text-sm font-medium text-[#8b8d9a] hover:text-[#ffffff] rounded-xl hover:bg-[#1e1e2a] transition-colors duration-200"
              >
                Cancel
              </button>
              <button
                onClick={handleDuplicateElevation}
                disabled={!duplicateName.trim() || !!elevations[duplicateName.trim()]}
                className="flex items-center gap-2 px-4 py-2.5 bg-gradient-to-r from-[#3b82f6] to-[#2563eb] hover:brightness-110 disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-semibold rounded-xl transition-colors duration-200"
              >
                <Copy className="w-4 h-4" />
                Duplicate
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Report Options Dialog */}
      <ReportOptionsDialog
        isOpen={showReportOptions}
        onClose={() => setShowReportOptions(false)}
        elevations={elevations}
        doors={doors}
        settings={settings}
        materials={materials}
        projectName={projectName}
        onSettingsUpdate={handleSettingsUpdate}
      />
    </div>
  );
}
