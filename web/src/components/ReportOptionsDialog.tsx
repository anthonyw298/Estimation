'use client';

import { useState, useCallback } from 'react';
import {
  X,
  FileSpreadsheet,
  FileText,
  ChevronDown,
  ChevronUp,
  Check,
  Copy,
} from 'lucide-react';
import type { ElevationData, DoorConfig, ProjectSettings, ExtraMaterial } from '@/types';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface ReportOptionsDialogProps {
  isOpen: boolean;
  onClose: () => void;
  elevations: Record<string, ElevationData>;
  doors: Record<string, DoorConfig[]>;
  settings: ProjectSettings;
  materials: Record<string, ExtraMaterial>;
  projectName: string;
  onSettingsUpdate: (newSettings: ProjectSettings) => Promise<void>;
}

// Material sections
const MATERIAL_SECTIONS = ['profiles', 'accessories', 'gaskets', 'doors', 'glass', 'fabrication'] as const;
const OTHER_SECTIONS = ['system_input', 'elevation_cost_summary', 'diagram'] as const;
const ALL_SECTIONS = [...MATERIAL_SECTIONS, ...OTHER_SECTIONS];

// Per-elevation column keys for material sections (matching Python's _build_elev_cols)
const ELEV_COLUMN_DEFS: Array<{ key: string; label: string; perElev?: boolean }> = [
  { key: 'description', label: 'Description' },
  { key: 'part_number', label: 'Part Number' },
  { key: 'total_quantity_required', label: 'Total Quantity Required' },
  { key: 'quantity_per_elevation', label: 'Quantity Per Elevation', perElev: true },
  { key: 'total_list_cost', label: 'Total List Cost' },
  { key: 'total_list_cost_per_elevation', label: 'Total List Cost Per Elevation', perElev: true },
  { key: 'discounted_total_list_cost', label: 'Discounted Total List Cost' },
  { key: 'discounted_total_list_cost_per_elevation', label: 'Discounted Total List Cost Per Elevation', perElev: true },
];

const SECTION_LABELS: Record<string, string> = {
  profiles: 'Profiles',
  accessories: 'Accessories',
  gaskets: 'Gaskets',
  doors: 'Doors',
  glass: 'Glass',
  fabrication: 'Fabrication',
  system_input: 'System Input',
  elevation_cost_summary: 'Elevation Cost Summary',
  diagram: 'Diagram',
};

// Summary cost overview keys
const COST_OVERVIEW_KEYS = ['additional_costs', 'markups', 'project_total', 'diagram'] as const;
const COST_OVERVIEW_LABELS: Record<string, string> = {
  additional_costs: 'Additional Costs',
  markups: 'Markups',
  project_total: 'Project Total',
  diagram: 'Pie Chart',
};

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export default function ReportOptionsDialog({
  isOpen,
  onClose,
  elevations,
  doors,
  settings,
  materials,
  projectName,
  onSettingsUpdate,
}: ReportOptionsDialogProps) {
  const elevationNames = Object.keys(elevations).sort();

  // State: per-elevation inclusion
  const [elevIncluded, setElevIncluded] = useState<Record<string, boolean>>(() =>
    Object.fromEntries(elevationNames.map(n => [n, true])),
  );

  // State: per-elevation sections
  const [elevSections, setElevSections] = useState<Record<string, Record<string, boolean>>>(() =>
    Object.fromEntries(
      elevationNames.map(n => [
        n,
        Object.fromEntries(ALL_SECTIONS.map(s => [s, true])),
      ]),
    ),
  );

  // State: per-elevation per-section column visibility
  // Structure: { elevName: { sectionKey: { colKey: boolean } } }
  const [elevColumns, setElevColumns] = useState<Record<string, Record<string, Record<string, boolean>>>>(() =>
    Object.fromEntries(
      elevationNames.map(n => [
        n,
        Object.fromEntries(
          MATERIAL_SECTIONS.map(s => [
            s,
            Object.fromEntries(ELEV_COLUMN_DEFS.map(c => [c.key, true])),
          ]),
        ),
      ]),
    ),
  );

  // State: summary
  const [summaryIncluded, setSummaryIncluded] = useState(true);
  const [summarySections, setSummarySections] = useState<Record<string, boolean>>(() =>
    Object.fromEntries(ALL_SECTIONS.map(s => [s, true])),
  );
  const [costOverview, setCostOverview] = useState<Record<string, boolean>>(() =>
    Object.fromEntries(COST_OVERVIEW_KEYS.map(k => [k, true])),
  );

  // State: collapsed panels
  const [expandedPanels, setExpandedPanels] = useState<Record<string, boolean>>({});

  // State: Elevation Summary Display (persisted to settings)
  const [showElevationNames, setShowElevationNames] = useState(settings.show_elevation_names ?? false);
  const [showElevationQuantity, setShowElevationQuantity] = useState(settings.show_elevation_quantity ?? false);
  const [showElevationDimensions, setShowElevationDimensions] = useState(settings.show_elevation_dimensions ?? false);
  const [showElevationSqft, setShowElevationSqft] = useState(settings.show_elevation_sqft ?? false);
  const [showElevationPerimeter, setShowElevationPerimeter] = useState(settings.show_elevation_perimeter ?? false);

  // Auto-save elevation summary display when toggles change
  const handleElevDisplayToggle = useCallback(
    (key: keyof ProjectSettings, value: boolean) => {
      const updatedSettings = { ...settings, [key]: value };
      onSettingsUpdate(updatedSettings);
    },
    [settings, onSettingsUpdate],
  );

  // State: export status
  const [exporting, setExporting] = useState(false);

  const togglePanel = (key: string) =>
    setExpandedPanels(prev => ({ ...prev, [key]: !prev[key] }));

  // Apply first elevation settings to all
  const applyToAll = useCallback(() => {
    if (elevationNames.length < 2) return;
    const firstSections = elevSections[elevationNames[0]];
    if (!firstSections) return;
    setElevSections(prev => {
      const next = { ...prev };
      for (const name of elevationNames.slice(1)) {
        next[name] = { ...firstSections };
      }
      return next;
    });
  }, [elevationNames, elevSections]);

  // Export handlers
  // Check if any included elevation has calculated data
  const hasCalculatedData = elevationNames.some(
    name => elevIncluded[name] && elevations[name]?.calculated_outputs && elevations[name].calculated_outputs!.length > 0
  );

  const handleExportExcel = useCallback(async () => {
    if (!hasCalculatedData) {
      alert('No calculated data to export. Please run "Calculate & Save" on at least one elevation first.');
      return;
    }
    setExporting(true);
    try {
      const { exportToExcel } = await import('@/lib/export');
      const filteredElevations: Record<string, ElevationData> = {};
      const filteredDoors: Record<string, DoorConfig[]> = {};
      for (const name of elevationNames) {
        if (elevIncluded[name]) {
          filteredElevations[name] = elevations[name];
          filteredDoors[name] = doors[name] || [];
        }
      }
      await exportToExcel(projectName, filteredElevations, filteredDoors, settings, materials);
      onClose();
    } catch (error) {
      console.error('Export failed:', error);
      alert('Export failed: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setExporting(false);
    }
  }, [elevationNames, elevIncluded, elevations, doors, settings, materials, projectName, onClose, hasCalculatedData]);

  const handleExportPDF = useCallback(async () => {
    if (!hasCalculatedData) {
      alert('No calculated data to export. Please run "Calculate & Save" on at least one elevation first.');
      return;
    }
    setExporting(true);
    try {
      const { exportToPdf } = await import('@/lib/pdf-export');
      const filteredElevations: Record<string, ElevationData> = {};
      const filteredDoors: Record<string, DoorConfig[]> = {};
      for (const name of elevationNames) {
        if (elevIncluded[name]) {
          filteredElevations[name] = elevations[name];
          filteredDoors[name] = doors[name] || [];
        }
      }
      await exportToPdf(projectName, filteredElevations, filteredDoors, settings, materials);
      onClose();
    } catch (error) {
      console.error('PDF export failed:', error);
      alert('PDF export failed: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setExporting(false);
    }
  }, [elevationNames, elevIncluded, elevations, doors, settings, materials, projectName, onClose, hasCalculatedData]);

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm"
      onClick={onClose}
    >
      <div
        className="bg-[#18181b] border border-[#27272a] rounded-2xl w-full max-w-2xl mx-4 max-h-[85vh] flex flex-col shadow-2xl animate-fade-in"
        onClick={e => e.stopPropagation()}
      >
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-[#27272a]">
          <h2 className="text-lg font-semibold text-white">Report Stock List</h2>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg hover:bg-[#27272a] text-zinc-400 hover:text-white transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Instructions */}
        <div className="px-6 py-3 text-xs text-zinc-500 border-b border-[#27272a]">
          Select which elevations, sections, and summary options to include in the exported report.
          {!hasCalculatedData && (
            <div className="mt-2 rounded bg-yellow-900/30 px-3 py-2 text-yellow-400 font-medium">
              No elevations have been calculated yet. Run &ldquo;Calculate &amp; Save&rdquo; on your elevations before exporting.
            </div>
          )}
        </div>

        {/* Content - scrollable */}
        <div className="flex-1 overflow-y-auto px-6 py-4 space-y-3">
          {/* Per-elevation panels */}
          {elevationNames.map(name => (
            <div key={name} className="border border-[#27272a] rounded-lg overflow-hidden">
              {/* Panel header */}
              <div className="flex items-center justify-between px-4 py-2.5 bg-[#111113]">
                <div className="flex items-center gap-3">
                  <input
                    type="checkbox"
                    checked={elevIncluded[name] ?? true}
                    onChange={e => setElevIncluded(prev => ({ ...prev, [name]: e.target.checked }))}
                    className="h-4 w-4 rounded border-zinc-600 bg-[#1c1c21] text-blue-500 accent-blue-500"
                  />
                  <span className="text-sm font-medium text-white">{name}</span>
                </div>
                <button
                  onClick={() => togglePanel(name)}
                  className="p-1 rounded hover:bg-[#27272a] text-zinc-400 transition-colors"
                >
                  {expandedPanels[name] ? (
                    <ChevronUp className="w-4 h-4" />
                  ) : (
                    <ChevronDown className="w-4 h-4" />
                  )}
                </button>
              </div>

              {/* Expanded content */}
              {expandedPanels[name] && (
                <div className="px-4 py-3 space-y-3 bg-[#0d0d0f]">
                  <p className="text-[10px] text-zinc-600 uppercase tracking-wider font-semibold">Sections & Columns</p>

                  {ALL_SECTIONS.map(section => {
                    const isMaterial = (MATERIAL_SECTIONS as readonly string[]).includes(section);
                    const sectionEnabled = elevSections[name]?.[section] ?? true;
                    const totalCount = elevations[name]?.total_count ?? 1;

                    return (
                      <div key={section}>
                        {/* Section checkbox */}
                        <label className="flex items-center gap-2 text-xs text-zinc-300 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={sectionEnabled}
                            onChange={e => {
                              setElevSections(prev => ({
                                ...prev,
                                [name]: { ...prev[name], [section]: e.target.checked },
                              }));
                            }}
                            className="h-3.5 w-3.5 rounded border-zinc-600 bg-[#1c1c21] text-blue-500 accent-blue-500"
                          />
                          <span className="font-medium">{SECTION_LABELS[section]}</span>
                        </label>

                        {/* Column toggles (only for material sections when enabled) */}
                        {isMaterial && sectionEnabled && (
                          <div className="ml-6 mt-1.5 mb-1 grid grid-cols-2 gap-1">
                            {ELEV_COLUMN_DEFS
                              .filter(col => !col.perElev || totalCount > 1)
                              .map(col => (
                                <label key={col.key} className="flex items-center gap-1.5 text-[10px] text-zinc-500 cursor-pointer">
                                  <input
                                    type="checkbox"
                                    checked={elevColumns[name]?.[section]?.[col.key] ?? true}
                                    onChange={e => {
                                      setElevColumns(prev => ({
                                        ...prev,
                                        [name]: {
                                          ...prev[name],
                                          [section]: {
                                            ...(prev[name]?.[section] ?? {}),
                                            [col.key]: e.target.checked,
                                          },
                                        },
                                      }));
                                    }}
                                    className="h-3 w-3 rounded border-zinc-700 bg-[#1c1c21] text-blue-500 accent-blue-500"
                                  />
                                  {col.label}
                                </label>
                              ))}
                          </div>
                        )}
                      </div>
                    );
                  })}

                  {elevations[name]?.total_count > 1 && (
                    <p className="text-[10px] text-zinc-600 italic mt-1">
                      &apos;Per Elevation&apos; columns only apply when elevation count &gt; 1
                    </p>
                  )}
                </div>
              )}
            </div>
          ))}

          {/* Summary panel */}
          <div className="border border-[#27272a] rounded-lg overflow-hidden">
            <div className="flex items-center justify-between px-4 py-2.5 bg-[#111113]">
              <div className="flex items-center gap-3">
                <input
                  type="checkbox"
                  checked={summaryIncluded}
                  onChange={e => setSummaryIncluded(e.target.checked)}
                  className="h-4 w-4 rounded border-zinc-600 bg-[#1c1c21] text-blue-500 accent-blue-500"
                />
                <span className="text-sm font-medium text-white">Summary</span>
              </div>
              <button
                onClick={() => togglePanel('summary')}
                className="p-1 rounded hover:bg-[#27272a] text-zinc-400 transition-colors"
              >
                {expandedPanels.summary ? (
                  <ChevronUp className="w-4 h-4" />
                ) : (
                  <ChevronDown className="w-4 h-4" />
                )}
              </button>
            </div>

            {expandedPanels.summary && (
              <div className="px-4 py-3 space-y-3 bg-[#0d0d0f]">
                <div>
                  <p className="text-[10px] text-zinc-600 uppercase tracking-wider font-semibold mb-1.5">Material Sections</p>
                  <div className="grid grid-cols-2 gap-1.5 sm:grid-cols-3">
                    {MATERIAL_SECTIONS.map(section => (
                      <label key={section} className="flex items-center gap-2 text-xs text-zinc-300 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={summarySections[section] ?? true}
                          onChange={e => setSummarySections(prev => ({ ...prev, [section]: e.target.checked }))}
                          className="h-3.5 w-3.5 rounded border-zinc-600 bg-[#1c1c21] text-blue-500 accent-blue-500"
                        />
                        {SECTION_LABELS[section]}
                      </label>
                    ))}
                  </div>
                </div>

                <div>
                  <p className="text-[10px] text-zinc-600 uppercase tracking-wider font-semibold mb-1.5">Cost Overview</p>
                  <div className="grid grid-cols-2 gap-1.5">
                    {COST_OVERVIEW_KEYS.map(key => (
                      <label key={key} className="flex items-center gap-2 text-xs text-zinc-300 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={costOverview[key] ?? true}
                          onChange={e => setCostOverview(prev => ({ ...prev, [key]: e.target.checked }))}
                          className="h-3.5 w-3.5 rounded border-zinc-600 bg-[#1c1c21] text-blue-500 accent-blue-500"
                        />
                        {COST_OVERVIEW_LABELS[key]}
                      </label>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Apply to all button */}
          {elevationNames.length > 1 && (
            <button
              onClick={applyToAll}
              className="flex items-center gap-2 text-xs text-zinc-400 hover:text-zinc-200 transition-colors"
            >
              <Copy className="w-3.5 h-3.5" />
              Apply first elevation&apos;s settings to all
            </button>
          )}

          {/* Elevation Summary Display */}
          <div className="border border-[#27272a] rounded-lg overflow-hidden">
            <div className="flex items-center justify-between px-4 py-2.5 bg-[#111113]">
              <span className="text-sm font-medium text-white">Elevation Summary</span>
              <button
                onClick={() => togglePanel('elevSummaryDisplay')}
                className="p-1 rounded hover:bg-[#27272a] text-zinc-400 transition-colors"
              >
                {expandedPanels.elevSummaryDisplay ? (
                  <ChevronUp className="w-4 h-4" />
                ) : (
                  <ChevronDown className="w-4 h-4" />
                )}
              </button>
            </div>

            {expandedPanels.elevSummaryDisplay && (
              <div className="px-4 py-3 space-y-2 bg-[#0d0d0f]">
                <p className="text-[10px] text-zinc-600">
                  Select which columns to include in the elevation summary table in exported reports.
                </p>
                <div className="grid grid-cols-2 gap-1.5 sm:grid-cols-3">
                  {([
                    { label: 'Elevation Name', key: 'show_elevation_names' as const, checked: showElevationNames, setter: setShowElevationNames },
                    { label: 'Quantity (EA)', key: 'show_elevation_quantity' as const, checked: showElevationQuantity, setter: setShowElevationQuantity },
                    { label: 'Dimensions', key: 'show_elevation_dimensions' as const, checked: showElevationDimensions, setter: setShowElevationDimensions },
                    { label: 'SQFT Total (SQFT)', key: 'show_elevation_sqft' as const, checked: showElevationSqft, setter: setShowElevationSqft },
                    { label: 'Perimeter FT Total (FT)', key: 'show_elevation_perimeter' as const, checked: showElevationPerimeter, setter: setShowElevationPerimeter },
                  ] as const).map(({ label, key, checked, setter }) => (
                    <label key={key} className="flex items-center gap-2 text-xs text-zinc-300 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={(e) => {
                          setter(e.target.checked);
                          handleElevDisplayToggle(key, e.target.checked);
                        }}
                        className="h-3.5 w-3.5 rounded border-zinc-600 bg-[#1c1c21] text-blue-500 accent-blue-500"
                      />
                      {label}
                    </label>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Footer - export buttons */}
        <div className="flex items-center justify-end gap-3 px-6 py-4 border-t border-[#27272a]">
          <button
            onClick={onClose}
            className="px-4 py-2 text-sm font-medium text-zinc-400 hover:text-white hover:bg-[#27272a] rounded-lg transition-colors"
          >
            Cancel
          </button>
          <button
            onClick={handleExportPDF}
            disabled={exporting}
            className="flex items-center gap-2 px-4 py-2 text-sm font-medium bg-red-600 hover:bg-red-700 text-white rounded-lg disabled:opacity-40 transition-colors"
          >
            <FileText className="w-4 h-4" />
            Export PDF
          </button>
          <button
            onClick={handleExportExcel}
            disabled={exporting}
            className="flex items-center gap-2 px-4 py-2 text-sm font-medium bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg disabled:opacity-40 transition-colors"
          >
            <FileSpreadsheet className="w-4 h-4" />
            Export Excel
          </button>
        </div>
      </div>
    </div>
  );
}
