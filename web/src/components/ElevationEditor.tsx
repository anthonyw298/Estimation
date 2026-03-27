'use client';

import { useState, useCallback, useMemo } from 'react';
import {
  ElevationData,
  DoorConfig,
  CalculatedOutput,
  ProjectSettings,
  ExtraMaterial,
  MaterialImpactDetails,
} from '@/types';
import { calculateYes45tuQuantities } from '@/lib/yes45tu';
import {
  buildDloGrid,
  calculateGlassMakeSize,
  DLO_EDGE_DEDUCTION,
  DLO_INTERIOR_DEDUCTION,
  DLO_SILL_DEDUCTION,
} from '@/lib/formulas';
import {
  getPriceByPart,
  getUnitPriceByPart,
  applyMaterialImpactInMemory,
  reverseMaterialImpact,
} from '@/lib/pricing';
import { calculate_door_info } from '@/lib/formulas';
import { Calculator, Save, ChevronDown, ChevronUp, DoorOpen, Layers, CheckCircle2, Eye, EyeOff, Table, AlertTriangle } from 'lucide-react';
import { PART_NUMBER_MAP } from '@/data/part-number';
import BayDiagram from './BayDiagram';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface ElevationEditorProps {
  projectName: string;
  elevationName: string;
  elevationData: ElevationData;
  doors: DoorConfig[];
  materials: Record<string, ExtraMaterial>;
  onSave: (
    name: string,
    data: ElevationData,
    doors: DoorConfig[],
    materials: Record<string, ExtraMaterial>,
  ) => void;
  settings: ProjectSettings;
}

const SYSTEM_TYPES = ['YES 45TU Front Set (OG)'];
const FINISHES = ['Clear', 'Black', 'Paint'];
const DOOR_SIZES = [
  "3' X 7'",
  "3' X 8'",
  "3' X 9'",
  "6' X 7'",
  "6' X 8'",
  "6' X 9'",
];
const STILE_TYPES = ['Narrow', 'Medium', 'Wide'];
const HARDWARE_OPTIONS = [
  'Concealed Closer',
  'Exit Devices',
  'Continuous Hinges',
  'Latch Lock w/ Lever Handle',
  'Lever Handle',
  'Electric Strike',
  'Extended Ladder Pull (B2B)',
  'Extended Ladder Pull (Single)',
];

// ---------------------------------------------------------------------------
// Material section ordering & labels (matching Python excel_generator)
// ---------------------------------------------------------------------------

const MATERIAL_SECTION_ORDER = ['profiles', 'accessories', 'gaskets', 'Glass', 'Fabrication', 'Doors'] as const;
const MATERIAL_SECTION_LABELS: Record<string, string> = {
  profiles: 'Profiles',
  accessories: 'Accessories',
  gaskets: 'Gaskets',
  Glass: 'Glass',
  Fabrication: 'Fabrication',
  Doors: 'Doors',
};

/** Types that receive the discount multiplier */
const DISCOUNTABLE_TYPES = new Set(['profiles', 'accessories', 'gaskets']);

/** Canonical column definitions matching Python's _build_elev_cols */
interface ColumnDef {
  key: string;
  label: string;
  perElev?: boolean; // only shown when totalCount > 1
}

const RESULTS_COLUMN_DEFS: ColumnDef[] = [
  { key: 'description', label: 'Description' },
  { key: 'part_number', label: 'Part Number' },
  { key: 'total_quantity_required', label: 'Total Quantity Required' },
  { key: 'quantity_per_elevation', label: 'Quantity Per Elevation', perElev: true },
  { key: 'total_list_cost', label: 'Total List Cost' },
  { key: 'total_list_cost_per_elevation', label: 'Total List Cost Per Elevation', perElev: true },
  { key: 'discounted_total_list_cost', label: 'Discounted Total List Cost' },
  { key: 'discounted_total_list_cost_per_elevation', label: 'Discounted Total List Cost Per Elevation', perElev: true },
];

// ---------------------------------------------------------------------------
// Shared input styling
// ---------------------------------------------------------------------------

const inputClass =
  'bg-[#0c0c12] border border-[#1e1e2a] text-white rounded-xl px-3.5 py-2.5 w-full focus:outline-none focus:ring-2 focus:ring-[#3b82f6]/20 focus:border-[#3b82f6] transition-colors duration-200 text-sm';
const inputInvalidClass =
  'bg-[#0c0c12] border border-red-500/50 text-white rounded-xl px-3.5 py-2.5 w-full focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-red-500 transition-colors duration-200 text-sm';
const selectClass =
  'bg-[#0c0c12] border border-[#1e1e2a] text-white rounded-xl px-3.5 py-2.5 w-full focus:outline-none focus:ring-2 focus:ring-[#3b82f6]/20 focus:border-[#3b82f6] transition-colors duration-200 appearance-none text-sm';
const labelClass = 'block text-sm font-medium text-[#ffffff] mb-1.5';
const cardClass =
  'bg-[#111118] border border-[#1e1e2a] rounded-2xl p-6 space-y-4 shadow-lg shadow-black/15';
const sectionTitleClass = 'text-lg font-semibold text-white tracking-tight';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function sumArray(arr: number[]): number {
  return arr.reduce((s, v) => s + v, 0);
}

function formatCurrency(value: number): string {
  return `$${value.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

function formatQuantity(
  qty: number | number[],
  type: string,
  unit?: string,
): string {
  if (type === 'profiles') {
    if (Array.isArray(qty)) {
      const total = sumArray(qty);
      return `${total.toFixed(2)} ft (${qty.length} pcs)`;
    }
    return `${Number(qty).toFixed(2)} ft`;
  }
  if (type === 'Glass') {
    return `${Number(qty).toFixed(2)} sqft`;
  }
  if (type === 'Fabrication') {
    return `${qty} joints`;
  }
  if (type === 'Calculations') {
    return `${Number(qty).toFixed(2)} ${unit ?? 'sqft'}`;
  }
  if (type === 'Doors') {
    return String(qty);
  }
  // accessories or unknown
  if (Array.isArray(qty)) {
    return String(sumArray(qty));
  }
  return String(qty);
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export default function ElevationEditor({
  projectName,
  elevationName,
  elevationData,
  doors: initialDoors,
  materials: initialMaterials,
  onSave,
  settings,
}: ElevationEditorProps) {
  // --- Door-only mode ---
  const [doorOnly, setDoorOnly] = useState(elevationData.door_only ?? false);

  // --- Form state ---
  const [systemType, setSystemType] = useState(
    elevationData.system_type || SYSTEM_TYPES[0],
  );
  const [finish, setFinish] = useState(elevationData.finish || 'Clear');
  const [totalCount, setTotalCount] = useState(elevationData.total_count || 0);
  const [openingWidth, setOpeningWidth] = useState(
    elevationData.opening_width_inches || 0,
  );
  const [openingHeight, setOpeningHeight] = useState(
    elevationData.opening_height_inches || 0,
  );
  const [baysWide, setBaysWide] = useState(elevationData.bays_wide || 0);
  const [baysTall, setBaysTall] = useState(elevationData.bays_tall || 0);
  const [customBayWidths, setCustomBayWidths] = useState<number[]>(
    () => {
      if (
        elevationData.custom_bay_widths &&
        elevationData.custom_bay_widths.length === elevationData.bays_wide
      ) {
        return elevationData.custom_bay_widths;
      }
      const w = elevationData.bays_wide || 1;
      const ow = elevationData.opening_width_inches || 0;
      return Array(w).fill(w > 0 ? ow / w : 0);
    },
  );
     const [customBayHeights, setCustomBayHeights] = useState<number[]>(
    () => {
      if (
        elevationData.custom_bay_heights &&
        elevationData.custom_bay_heights.length === elevationData.bays_tall
      ) {
        return elevationData.custom_bay_heights;
      }
      const h = elevationData.bays_tall || 1;
      const oh = elevationData.opening_height_inches || 0;
      return Array(h).fill(h > 0 ? oh / h : 0);
    },
  );

  // Track which bay indices the user has manually edited (for "Default Rest")
  const [editedBayWidths, setEditedBayWidths] = useState<Set<number>>(new Set());
  const [editedBayHeights, setEditedBayHeights] = useState<Set<number>>(new Set());

  // Glass & Fabrication pricing is controlled from the Pricing tab (settings),
  // not per-elevation. Read from settings (fallback to elevation data for legacy).
  const glassPerSqft = settings.glass_per_sqft ?? elevationData.glass_per_sqft ?? 10.5;
  const fabCostPerJoint = settings.fabrication_cost_per_joint ?? elevationData.fabrication_cost_per_joint ?? 15.0;

  // --- Doors ---
  const [doors, setDoors] = useState<DoorConfig[]>(initialDoors);

  // --- Field Costs (per-elevation) ---
  const [installationLaborHours, setInstallationLaborHours] = useState(
    elevationData.installation_labor_hours ?? 0,
  );
  const [sealantJoints, setSealantJoints] = useState(
    elevationData.sealant_joints ?? 0,
  );
  const [breakMetalSelections, setBreakMetalSelections] = useState<string[]>(
    elevationData.break_metal_selections ?? [],
  );
  const [fieldCostExpanded, setFieldCostExpanded] = useState(
    (elevationData.installation_labor_hours ?? 0) > 0 ||
    (elevationData.sealant_joints ?? 0) > 0 ||
    (elevationData.break_metal_selections ?? []).length > 0,
  );

  // --- Results ---
  const [results, setResults] = useState<CalculatedOutput[] | null>(
    elevationData.calculated_outputs ?? null,
  );
  const [materialImpacts, setMaterialImpacts] = useState<
    MaterialImpactDetails[]
  >(elevationData.material_impacts ?? []);
  const [isCalculating, setIsCalculating] = useState(false);

  // --- Save success indicator ---
  const [justSaved, setJustSaved] = useState(false);

  // --- Validation: only show red highlights after first save attempt ---
  const [saveAttempted, setSaveAttempted] = useState(false);

  // --- Column visibility for results table ---
  const [visibleColumns, setVisibleColumns] = useState<Set<string>>(
    () => new Set(['description', 'part_number', 'total_quantity_required', 'total_list_cost', 'discounted_total_list_cost']),
  );
  const [showColumnPicker, setShowColumnPicker] = useState(false);

  const toggleColumnVisibility = useCallback((key: string) => {
    setVisibleColumns((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  }, []);

  // --- Results table expand/collapse ---
  const [resultsTableExpanded, setResultsTableExpanded] = useState(false);

  // --- Section collapse (results/materialImpact collapsed by default) ---
  const [collapsedSections, setCollapsedSections] = useState<
    Record<string, boolean>
  >({ results: true, materialImpact: true });
  const toggleSection = (key: string) =>
    setCollapsedSections((prev) => ({ ...prev, [key]: !prev[key] }));

  // ---------------------------------------------------------------------------
  // Bay width logic
  // ---------------------------------------------------------------------------

  const bayWidthSum = useMemo(() => sumArray(customBayWidths), [customBayWidths]);
  const bayWidthMismatch = useMemo(
    () => baysWide > 1 && Math.abs(bayWidthSum - openingWidth) > 0.01,
    [baysWide, bayWidthSum, openingWidth],
  );

  const handleBaysWideChange = useCallback(
    (newBaysWide: number) => {
      setBaysWide(newBaysWide);
      setEditedBayWidths(new Set());
      const count = Math.max(1, newBaysWide);
      if (openingWidth > 0) {
        setCustomBayWidths(
          Array(count).fill(
            Math.round((openingWidth / count) * 100) / 100,
          ),
        );
      } else {
        setCustomBayWidths(Array(count).fill(0));
      }
    },
    [openingWidth],
  );

  const handleOpeningWidthChange = useCallback(
    (newWidth: number) => {
      setOpeningWidth(newWidth);
      if (baysWide > 1 && newWidth > 0) {
        if (editedBayWidths.size === 0) {
          // No manual edits — redistribute equally
          setCustomBayWidths(
            Array(baysWide).fill(Math.round((newWidth / baysWide) * 100) / 100),
          );
        } else if (editedBayWidths.size < baysWide) {
          // Some manual edits — keep those, redistribute rest
          setCustomBayWidths((prev) => {
            const editedSum = prev.reduce(
              (sum, w, i) => (editedBayWidths.has(i) ? sum + w : sum), 0,
            );
            const remaining = newWidth - editedSum;
            const unedited = baysWide - editedBayWidths.size;
            const each = Math.round((remaining / unedited) * 100) / 100;
            return prev.map((w, i) => (editedBayWidths.has(i) ? w : each));
          });
        }
        // All edited — don't touch, mismatch warning will show
      }
    },
    [baysWide, editedBayWidths],
  );

  const handleCustomBayWidthChange = useCallback(
    (index: number, value: number) => {
      setCustomBayWidths((prev) => {
        const next = [...prev];
        next[index] = value;
        return next;
      });
      // Mark as edited when value > 0, remove when cleared to 0
      setEditedBayWidths((prev) => {
        const next = new Set(prev);
        if (value > 0) next.add(index); else next.delete(index);
        return next;
      });
    },
    [],
  );

  /** Reset all bay widths to equal split. */
  const handleDefaultAllWidths = useCallback(() => {
    const each = Math.round((openingWidth / baysWide) * 100) / 100;
    setCustomBayWidths(Array(baysWide).fill(each));
    setEditedBayWidths(new Set());
  }, [openingWidth, baysWide]);

  /** Keep edited bays, distribute remaining to un-edited bays. */
  const handleDefaultRestWidths = useCallback(() => {
    if (editedBayWidths.size === 0 || editedBayWidths.size >= baysWide) {
      // Nothing edited or all edited — just do equal split
      handleDefaultAllWidths();
      return;
    }
    const editedSum = customBayWidths.reduce(
      (sum, w, i) => (editedBayWidths.has(i) ? sum + w : sum),
      0,
    );
    const remaining = openingWidth - editedSum;
    const unedited = baysWide - editedBayWidths.size;
    const each = Math.round((remaining / unedited) * 100) / 100;
    setCustomBayWidths((prev) =>
      prev.map((w, i) => (editedBayWidths.has(i) ? w : each)),
    );
  }, [editedBayWidths, customBayWidths, openingWidth, baysWide, handleDefaultAllWidths]);

  // ---------------------------------------------------------------------------
  // Bay height logic
  // ---------------------------------------------------------------------------

  const bayHeightSum = useMemo(() => sumArray(customBayHeights), [customBayHeights]);
  const bayHeightMismatch = useMemo(
    () => baysTall > 1 && Math.abs(bayHeightSum - openingHeight) > 0.01,
    [baysTall, bayHeightSum, openingHeight],
  );

  const handleBaysTallChange = useCallback(
    (newBaysTall: number) => {
      setBaysTall(newBaysTall);
      setEditedBayHeights(new Set());
      const count = Math.max(1, newBaysTall);
      if (openingHeight > 0) {
        setCustomBayHeights(
          Array(count).fill(
            Math.round((openingHeight / count) * 100) / 100,
          ),
        );
      } else {
        setCustomBayHeights(Array(count).fill(0));
      }
    },
    [openingHeight],
  );

  const handleOpeningHeightChange = useCallback(
    (newHeight: number) => {
      setOpeningHeight(newHeight);
      if (baysTall > 1 && newHeight > 0) {
        if (editedBayHeights.size === 0) {
          setCustomBayHeights(
            Array(baysTall).fill(Math.round((newHeight / baysTall) * 100) / 100),
          );
        } else if (editedBayHeights.size < baysTall) {
          setCustomBayHeights((prev) => {
            const editedSum = prev.reduce(
              (sum, h, i) => (editedBayHeights.has(i) ? sum + h : sum), 0,
            );
            const remaining = newHeight - editedSum;
            const unedited = baysTall - editedBayHeights.size;
            const each = Math.round((remaining / unedited) * 100) / 100;
            return prev.map((h, i) => (editedBayHeights.has(i) ? h : each));
          });
        }
      }
    },
    [baysTall, editedBayHeights],
  );

  const handleCustomBayHeightChange = useCallback(
    (index: number, value: number) => {
      setCustomBayHeights((prev) => {
        const next = [...prev];
        next[index] = value;
        return next;
      });
      setEditedBayHeights((prev) => {
        const next = new Set(prev);
        if (value > 0) next.add(index); else next.delete(index);
        return next;
      });
    },
    [],
  );

  /** Reset all bay heights to equal split. */
  const handleDefaultAllHeights = useCallback(() => {
    const each = Math.round((openingHeight / baysTall) * 100) / 100;
    setCustomBayHeights(Array(baysTall).fill(each));
    setEditedBayHeights(new Set());
  }, [openingHeight, baysTall]);

  /** Keep edited bays, distribute remaining to un-edited bays. */
  const handleDefaultRestHeights = useCallback(() => {
    if (editedBayHeights.size === 0 || editedBayHeights.size >= baysTall) {
      handleDefaultAllHeights();
      return;
    }
    const editedSum = customBayHeights.reduce(
      (sum, h, i) => (editedBayHeights.has(i) ? sum + h : sum),
      0,
    );
    const remaining = openingHeight - editedSum;
    const unedited = baysTall - editedBayHeights.size;
    const each = Math.round((remaining / unedited) * 100) / 100;
    setCustomBayHeights((prev) =>
      prev.map((h, i) => (editedBayHeights.has(i) ? h : each)),
    );
  }, [editedBayHeights, customBayHeights, openingHeight, baysTall, handleDefaultAllHeights]);

  // ---------------------------------------------------------------------------
  // Door management
  // ---------------------------------------------------------------------------

  const addDoor = useCallback(() => {
    setDoors((prev) => [
      ...prev,
      {
        size: "3' X 7'",
        count: 1,
        stile: 'Narrow',
        hardware: Object.fromEntries(HARDWARE_OPTIONS.map((h) => [h, false])),
      },
    ]);
  }, []);

  const removeDoor = useCallback((index: number) => {
    setDoors((prev) => prev.filter((_, i) => i !== index));
  }, []);

  const updateDoor = useCallback(
    (index: number, field: keyof DoorConfig, value: unknown) => {
      setDoors((prev) => {
        const next = [...prev];
        next[index] = { ...next[index], [field]: value };
        return next;
      });
    },
    [],
  );

  const updateDoorHardware = useCallback(
    (doorIndex: number, hw: string, checked: boolean) => {
      setDoors((prev) => {
        const next = [...prev];
        next[doorIndex] = {
          ...next[doorIndex],
          hardware: { ...next[doorIndex].hardware, [hw]: checked },
        };
        return next;
      });
    },
    [],
  );

  // ---------------------------------------------------------------------------
  // Invalidate stale results when inputs change
  // ---------------------------------------------------------------------------

  // ---------------------------------------------------------------------------
  // Update (calculate + save)
  // Whether this elevation has been saved before (has calculated_outputs)
  const isExisting = !!(elevationData.calculated_outputs && elevationData.calculated_outputs.length > 0);

  // ---------------------------------------------------------------------------
  // Validation: collect missing required fields
  // ---------------------------------------------------------------------------
  const missingFields = useMemo(() => {
    if (doorOnly) return [];
    const missing: string[] = [];
    if (openingWidth <= 0) missing.push('Opening Width');
    if (openingHeight <= 0) missing.push('Opening Height');
    if (baysWide <= 0) missing.push('Bays Wide');
    if (baysTall <= 0) missing.push('Bays Tall');
    if (totalCount <= 0) missing.push('Total Count');
    return missing;
  }, [doorOnly, openingWidth, openingHeight, baysWide, baysTall, totalCount]);

  const isFormValid = doorOnly ? doors.length > 0 : missingFields.length === 0;

  const handleCalculate = useCallback(() => {
    setSaveAttempted(true);
    // Safety net: block calculation if required fields are missing
    if (!doorOnly && missingFields.length > 0) {
      return;
    }
    if (doorOnly && doors.length === 0) {
      return;
    }

    setIsCalculating(true);

    try {
      const pricedOutputs: CalculatedOutput[] = [];
      const impacts: MaterialImpactDetails[] = [];

      // Deep-clone materials so we can track impact without mutating props
      const materialsClone: Record<string, ExtraMaterial> = {};
      for (const [k, v] of Object.entries(initialMaterials)) {
        materialsClone[k] = {
          quantity: v.quantity,
          length_pieces: [...v.length_pieces],
        };
      }

      // CRITICAL: Reverse old material impacts before recalculating
      // Without this, leftovers from previous calc cause $0 pricing on re-calc
      if (elevationData.material_impacts && elevationData.material_impacts.length > 0) {
        reverseMaterialImpact(elevationData.material_impacts, materialsClone);
      }

      if (doorOnly) {
        // ---- Door-only mode: only calculate doors ----
        if (doors.length === 0) {
          alert('Please add at least one door for door-only mode.');
          setIsCalculating(false);
          return;
        }

        const doorItems = calculate_door_info(doors, finish, 1);
        for (const doorItem of doorItems) {
          pricedOutputs.push({
            description: doorItem.description,
            quantity: doorItem.quantity,
            part_number: doorItem.part_number,
            type: doorItem.type,
            price: doorItem.price * doorItem.quantity,
            manual: true,
            hardware: doorItem.hardware,
            Style: doorItem.Style,
          });
        }
      } else {
        // ---- Standard elevation mode ----
        // 1. Run the YES 45TU quantity calculations
        const rawOutputs = calculateYes45tuQuantities(
          baysWide,
          baysTall,
          totalCount,
          openingWidth,
          openingHeight,
          doors,
          baysWide > 1 ? customBayWidths : undefined,
          baysTall > 1 ? customBayHeights : undefined,
          glassPerSqft,
          fabCostPerJoint,
        );

        for (const output of rawOutputs) {
          if (output.manual) {
            // Manual items: Glass, Fabrication, Calculations
            let totalPrice = 0;
            if (output.type === 'Glass' && typeof output.quantity === 'number') {
              totalPrice = output.quantity * (output.price ?? glassPerSqft);
            } else if (
              output.type === 'Fabrication' &&
              typeof output.quantity === 'number'
            ) {
              totalPrice = output.quantity * (output.price ?? fabCostPerJoint);
            }

            pricedOutputs.push({
              ...output,
              price:
                output.type === 'Calculations'
                  ? undefined
                  : totalPrice,
            });
          } else {
            // Standard parts: get pricing + material impact
            // Gaskets need group=true for proper length-based pricing (sold in 500ft rolls)
            const GASKET_PARTS = new Set(['E2-0052', 'E2-0053', 'E2-0065']);
            const isGasket =
              (output.description?.toLowerCase().includes('gasket') ?? false) ||
              GASKET_PARTS.has(output.part_number);
            const isProfile = output.type === 'profiles';
            const useGroup = isProfile || isGasket;

            // 1. Get theoretical price (summary=true) for display – always shows cost
            //    This matches the Python excel_generator summary sheet behavior.
            const [summaryPrice, unitType] = getPriceByPart(
              output.part_number,
              output.quantity,
              finish,
              null,
              true, // summary=true: ignores inventory, always returns purchase cost
              useGroup,
              output.description,
            );

            // 2. Get inventory-aware price + material impact for stock tracking
            const [, , impact] = getPriceByPart(
              output.part_number,
              output.quantity,
              finish,
              materialsClone,
              false, // summary=false: uses inventory, returns $0 for leftover-fulfilled items
              useGroup,
              output.description,
            );

            // Apply impact to clone for inventory tracking
            if (impact) {
              applyMaterialImpactInMemory(materialsClone, impact);
              impacts.push(impact);
            }

            pricedOutputs.push({
              ...output,
              price: summaryPrice ?? 0,
              unit: unitType ?? undefined,
            });
          }
        }

        // Door items for standard mode
        const doorItems = calculate_door_info(doors, finish, totalCount);
        for (const doorItem of doorItems) {
          pricedOutputs.push({
            description: doorItem.description,
            quantity: doorItem.quantity,
            part_number: doorItem.part_number,
            type: doorItem.type,
            price: doorItem.price * doorItem.quantity,
            manual: true,
            hardware: doorItem.hardware,
            Style: doorItem.Style,
          });
        }
      }

      setResults(pricedOutputs);
      setMaterialImpacts(impacts);

      // ------------------------------------------------------------------
      // Single-elevation outputs (count=1, no residual) for true per-elev cost
      // ------------------------------------------------------------------
      let singleElevOutputs: CalculatedOutput[] | undefined;
      if (!doorOnly && totalCount > 1) {
        singleElevOutputs = [];
        const GASKET_PARTS_SET = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

        // 1. Run formulas with count=1
        const singleRawOutputs = calculateYes45tuQuantities(
          baysWide,
          baysTall,
          1,
          openingWidth,
          openingHeight,
          doors,
          baysWide > 1 ? customBayWidths : undefined,
          baysTall > 1 ? customBayHeights : undefined,
          glassPerSqft,
          fabCostPerJoint,
        );

        // 2. Price each output (summary mode, no inventory)
        for (const output of singleRawOutputs) {
          if (output.manual) {
            let totalPrice = 0;
            if (output.type === 'Glass' && typeof output.quantity === 'number') {
              totalPrice = output.quantity * (output.price ?? glassPerSqft);
            } else if (output.type === 'Fabrication' && typeof output.quantity === 'number') {
              totalPrice = output.quantity * (output.price ?? fabCostPerJoint);
            }
            singleElevOutputs.push({
              ...output,
              price: output.type === 'Calculations' ? undefined : totalPrice,
            });
          } else {
            const isGasket =
              (output.description?.toLowerCase().includes('gasket') ?? false) ||
              GASKET_PARTS_SET.has(output.part_number);
            const isProfile = output.type === 'profiles';
            const useGroup = isProfile || isGasket;
            const [price, unitType] = getPriceByPart(
              output.part_number,
              output.quantity,
              finish,
              null,
              true,
              useGroup,
              output.description,
            );
            singleElevOutputs.push({
              ...output,
              price: price ?? 0,
              unit: unitType ?? undefined,
            });
          }
        }

        // 3. Door items for count=1
        const singleDoorItems = calculate_door_info(doors, finish, 1);
        for (const doorItem of singleDoorItems) {
          singleElevOutputs.push({
            description: doorItem.description,
            quantity: doorItem.quantity,
            part_number: doorItem.part_number,
            type: doorItem.type,
            price: doorItem.price * doorItem.quantity,
            manual: true,
            hardware: doorItem.hardware,
            Style: doorItem.Style,
          });
        }
      }

      // Build updated elevation data and call onSave
      // Field cost fields (preserved across saves)
      const fieldCostFields = {
        installation_labor_hours: installationLaborHours || undefined,
        sealant_joints: sealantJoints || undefined,
        break_metal_selections: breakMetalSelections.length > 0 ? breakMetalSelections : undefined,
      };

      const updatedData: ElevationData = doorOnly
        ? {
            system_type: 'Other',
            finish,
            opening_width_inches: doors[0]
              ? parseInt(doors[0].size.split('X')[0].replace("'", '').trim()) * 12
              : 0,
            opening_height_inches: doors[0]
              ? parseInt(doors[0].size.split('X')[1].replace("'", '').trim()) * 12
              : 0,
            bays_wide: 0,
            bays_tall: 0,
            total_count: 1,
            door_only: true,
            calculated_outputs: pricedOutputs,
            material_impacts: impacts,
            ...fieldCostFields,
          }
        : {
            system_type: systemType,
            finish,
            opening_width_inches: openingWidth,
            opening_height_inches: openingHeight,
            bays_wide: baysWide,
            bays_tall: baysTall,
            total_count: totalCount,
            custom_bay_widths: baysWide > 1 ? customBayWidths : undefined,
            custom_bay_heights: baysTall > 1 ? customBayHeights : undefined,
            glass_per_sqft: glassPerSqft,
            fabrication_cost_per_joint: fabCostPerJoint,
            calculated_outputs: pricedOutputs,
            single_elevation_outputs: singleElevOutputs,
            material_impacts: impacts,
            door_only: false,
            ...fieldCostFields,
          };

      onSave(elevationName, updatedData, doors, materialsClone);

      // Show success indicator
      setJustSaved(true);
      setTimeout(() => setJustSaved(false), 3000);
    } finally {
      setIsCalculating(false);
    }
  }, [
    doorOnly,
    baysWide,
    baysTall,
    totalCount,
    openingWidth,
    openingHeight,
    doors,
    customBayWidths,
    customBayHeights,
    glassPerSqft,
    fabCostPerJoint,
    finish,
    systemType,
    elevationName,
    initialMaterials,
    onSave,
    elevationData,
    missingFields,
    isFormValid,
    installationLaborHours,
    sealantJoints,
    breakMetalSelections,
  ]);

  // ---------------------------------------------------------------------------
  // Grand total
  // ---------------------------------------------------------------------------

  const grandTotal = useMemo(() => {
    if (!results) return 0;
    return results.reduce((sum, r) => {
      if (r.type === 'Calculations') return sum; // info-only row
      return sum + (r.price ?? 0);
    }, 0);
  }, [results]);

  // ---------------------------------------------------------------------------
  // Discount multiplier (same logic as CostSummary.tsx)
  // ---------------------------------------------------------------------------

  const discountMultiplier = useMemo(() => {
    if (settings.discount_multiplier != null) return settings.discount_multiplier;
    const threshold = settings.discount_threshold ?? 50000;
    const lowMult = settings.discount_multiplier_low ?? 0.614;
    const highMult = settings.discount_multiplier_high ?? 0.572;
    // Use grandTotal as a rough proxy — exact match requires project-wide total
    // but for per-elevation display this is reasonable
    return grandTotal < threshold ? lowMult : highMult;
  }, [settings, grandTotal]);

  // ---------------------------------------------------------------------------
  // Grouped results for table display
  // ---------------------------------------------------------------------------

  const groupedResults = useMemo(() => {
    if (!results) return [];

    // Classify each item into a section
    const groups: Record<string, CalculatedOutput[]> = {};
    const profileParts = new Set(Object.keys(PART_NUMBER_MAP['profiles'] ?? {}));
    const accessoryParts = new Set(Object.keys(PART_NUMBER_MAP['accessories'] ?? {}));
    const GASKET_PARTS = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

    for (const item of results) {
      if (item.type === 'Calculations') continue; // skip info-only rows

      let section = item.type; // default from calculation engine

      // Refine classification to match Python sections
      if (section === 'profiles' || profileParts.has(item.part_number)) {
        // Check if gasket (gaskets are in the profiles section by type but have specific part numbers)
        const isGasket =
          GASKET_PARTS.has(item.part_number) ||
          (item.description?.toLowerCase().includes('gasket') ?? false);
        section = isGasket ? 'gaskets' : 'profiles';
      } else if (section === 'accessories' || accessoryParts.has(item.part_number)) {
        section = 'accessories';
      }
      // Glass, Fabrication, Doors stay as-is

      if (!groups[section]) groups[section] = [];
      groups[section].push(item);
    }

    // Return in canonical order, filtering empty groups
    return MATERIAL_SECTION_ORDER
      .filter((s) => groups[s] && groups[s].length > 0)
      .map((s) => ({
        section: s,
        label: MATERIAL_SECTION_LABELS[s] ?? s,
        items: groups[s],
        isDiscountable: DISCOUNTABLE_TYPES.has(s),
      }));
  }, [results]);

  // Single-elevation lookup for true per-elev values (count=1, no residual)
  const singleElevMap = useMemo(() => {
    const map = new Map<string, { price: number; quantity: number | number[] }>();
    const singleOutputs = elevationData.single_elevation_outputs;
    if (!singleOutputs || totalCount <= 1) return map;

    const profileParts = new Set(Object.keys(PART_NUMBER_MAP['profiles'] ?? {}));
    const accessoryParts = new Set(Object.keys(PART_NUMBER_MAP['accessories'] ?? {}));
    const GASKET_PARTS_S = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

    for (const item of singleOutputs) {
      if (item.type === 'Calculations') continue;
      let section = item.type;
      if (section === 'profiles' || profileParts.has(item.part_number)) {
        const isGasket = GASKET_PARTS_S.has(item.part_number) || (item.description?.toLowerCase().includes('gasket') ?? false);
        section = isGasket ? 'gaskets' : 'profiles';
      } else if (section === 'accessories' || accessoryParts.has(item.part_number)) {
        section = 'accessories';
      }
      const key = `${section}|${item.description}|${item.part_number}`;
      map.set(key, { price: item.price ?? 0, quantity: item.quantity });
    }
    return map;
  }, [elevationData.single_elevation_outputs, totalCount]);

  // Active columns based on visibility and totalCount
  const activeColumns = useMemo(() => {
    return RESULTS_COLUMN_DEFS.filter((col) => {
      if (!visibleColumns.has(col.key)) return false;
      if (col.perElev && totalCount <= 1) return false;
      return true;
    });
  }, [visibleColumns, totalCount]);

  // ---------------------------------------------------------------------------
  // Render helpers
  // ---------------------------------------------------------------------------

  const SectionHeader = ({
    sectionKey,
    title,
    icon,
  }: {
    sectionKey: string;
    title: string;
    icon?: React.ReactNode;
  }) => (
    <button
      type="button"
      className="flex w-full items-center justify-between"
      onClick={() => toggleSection(sectionKey)}
    >
      <div className="flex items-center gap-2">
        {icon}
        <h3 className={sectionTitleClass}>{title}</h3>
      </div>
      {collapsedSections[sectionKey] ? (
        <ChevronDown className="h-5 w-5 text-[#ffffff]" />
      ) : (
        <ChevronUp className="h-5 w-5 text-[#ffffff]" />
      )}
    </button>
  );

  // ---------------------------------------------------------------------------
  // JSX
  // ---------------------------------------------------------------------------

  return (
    <div className="mx-auto max-w-5xl space-y-6 pb-12">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <p className="text-xs text-[#ffffff] font-medium tracking-wide uppercase">{projectName}</p>
          <h2 className="text-2xl font-bold text-white tracking-tight">{elevationName}</h2>
        </div>
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* Mode Toggle: Elevation vs Door Only */}
      {/* ------------------------------------------------------------------ */}
      <div className="flex items-center gap-3">
        <div className="flex items-center bg-[#0a0a10] border border-[#1e1e2a] rounded-xl p-0.5">
          <button
            type="button"
            onClick={() => setDoorOnly(false)}
            className={`flex items-center gap-2 px-4 py-2.5 text-sm font-medium rounded-lg transition-colors duration-200 ${
              !doorOnly
                ? 'bg-gradient-to-r from-[#3b82f6] to-[#2563eb] text-white'
                : 'text-[#ffffff] hover:text-[#ffffff]'
            }`}
          >
            <Layers className="w-4 h-4" />
            Elevation
          </button>
          <button
            type="button"
            onClick={() => setDoorOnly(true)}
            className={`flex items-center gap-2 px-4 py-2.5 text-sm font-medium rounded-lg transition-colors duration-200 ${
              doorOnly
                ? 'bg-gradient-to-r from-[#3b82f6] to-[#2563eb] text-white'
                : 'text-[#ffffff] hover:text-[#ffffff]'
            }`}
          >
            <DoorOpen className="w-4 h-4" />
            Door Only
          </button>
        </div>
        {doorOnly && (
          <span className="text-xs text-[#ffffff] ml-1">
            Door-only mode: no system, bays, or glass — just doors.
          </span>
        )}
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* 1. System Configuration (hidden in door-only mode) */}
      {/* ------------------------------------------------------------------ */}
      {!doorOnly && (
        <div className={cardClass}>
          <SectionHeader sectionKey="system" title="System Configuration" />
          {!collapsedSections.system && (
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
              <div>
                <label className={labelClass}>System Type</label>
                <select
                  className={selectClass}
                  value={systemType}
                  onChange={(e) => setSystemType(e.target.value)}
                >
                  {SYSTEM_TYPES.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className={labelClass}>Finish</label>
                <select
                  className={selectClass}
                  value={finish}
                  onChange={(e) => setFinish(e.target.value)}
                >
                  {FINISHES.map((f) => (
                    <option key={f} value={f}>
                      {f}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className={labelClass}>Total Count {saveAttempted && totalCount <= 0 && <span className="text-red-400">*</span>}</label>
                <input
                  type="number"
                  className={saveAttempted && totalCount <= 0 ? inputInvalidClass : inputClass}
                  min={1}
                  value={totalCount || ''}
                  onChange={(e) =>
                    setTotalCount(parseInt(e.target.value, 10) || 0)
                  }
                  onBlur={() => { if (totalCount < 1) setTotalCount(1); }}
                  placeholder="1"
                />
              </div>
            </div>
          )}
        </div>
      )}

      {/* Finish selector for door-only mode */}
      {doorOnly && (
        <div className={cardClass}>
          <h3 className="text-lg font-semibold text-white tracking-tight">Door Finish</h3>
          <div className="max-w-xs">
            <label className={labelClass}>Finish</label>
            <select
              className={selectClass}
              value={finish}
              onChange={(e) => setFinish(e.target.value)}
            >
              {FINISHES.map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
          </div>
        </div>
      )}

      {/* ------------------------------------------------------------------ */}
      {/* 2. Opening Dimensions (hidden in door-only mode) */}
      {/* ------------------------------------------------------------------ */}
      {!doorOnly && (
        <div className={cardClass}>
          <SectionHeader sectionKey="opening" title="Opening Dimensions" />
          {!collapsedSections.opening && (
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
              <div>
                <label className={labelClass}>Opening Width (inches) {saveAttempted && openingWidth <= 0 && <span className="text-red-400">*</span>}</label>
                <input
                  type="number"
                  className={saveAttempted && openingWidth <= 0 ? inputInvalidClass : inputClass}
                  min={0}
                  step="0.01"
                  value={openingWidth || ''}
                  onChange={(e) =>
                    handleOpeningWidthChange(parseFloat(e.target.value) || 0)
                  }
                  placeholder="e.g. 120"
                />
              </div>
              <div>
                <label className={labelClass}>Opening Height (inches) {saveAttempted && openingHeight <= 0 && <span className="text-red-400">*</span>}</label>
                <input
                  type="number"
                  className={saveAttempted && openingHeight <= 0 ? inputInvalidClass : inputClass}
                  min={0}
                  step="0.01"
                  value={openingHeight || ''}
                  onChange={(e) =>
                    handleOpeningHeightChange(parseFloat(e.target.value) || 0)
                  }
                  placeholder="e.g. 96"
                />
              </div>
            </div>
          )}
        </div>
      )}

      {/* ------------------------------------------------------------------ */}
      {/* 3. Bay Configuration (hidden in door-only mode) */}
      {/* ------------------------------------------------------------------ */}
      {!doorOnly && (
        <div className={cardClass}>
          <SectionHeader sectionKey="bay" title="Bay Configuration" />
          {!collapsedSections.bay && (
            <>
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                <div>
                  <label className={labelClass}>Bays Wide {saveAttempted && baysWide <= 0 && <span className="text-red-400">*</span>}</label>
                  <input
                    type="number"
                    className={saveAttempted && baysWide <= 0 ? inputInvalidClass : inputClass}
                    min={1}
                    value={baysWide || ''}
                    onChange={(e) =>
                      handleBaysWideChange(parseInt(e.target.value, 10) || 0)
                    }
                    onBlur={() => { if (baysWide < 1) handleBaysWideChange(1); }}
                    placeholder="1"
                  />
                </div>
                <div>
                  <label className={labelClass}>Bays Tall {saveAttempted && baysTall <= 0 && <span className="text-red-400">*</span>}</label>
                  <input
                    type="number"
                    className={saveAttempted && baysTall <= 0 ? inputInvalidClass : inputClass}
                    min={1}
                    value={baysTall || ''}
                    onChange={(e) =>
                      handleBaysTallChange(parseInt(e.target.value, 10) || 0)
                    }
                    onBlur={() => { if (baysTall < 1) handleBaysTallChange(1); }}
                    placeholder="1"
                  />
                </div>
              </div>

              {baysWide > 1 && (
                <div className="mt-3 space-y-3">
                  <div className="flex items-center gap-3">
                    <p className="text-sm font-medium text-[#ffffff]">
                      Custom Bay Widths (inches)
                    </p>
                    <button
                      type="button"
                      onClick={handleDefaultAllWidths}
                      className="rounded bg-[#1e1e2a] px-2.5 py-1 text-xs font-medium text-[#ffffff]/70 transition-colors hover:bg-[#2a2a3a] hover:text-[#ffffff]"
                    >
                      Default All
                    </button>
                    {editedBayWidths.size > 0 && editedBayWidths.size < baysWide && (
                      <button
                        type="button"
                        onClick={handleDefaultRestWidths}
                        className="rounded bg-blue-600/20 px-2.5 py-1 text-xs font-medium text-blue-400 transition-colors hover:bg-blue-600/30"
                      >
                        Default Rest
                      </button>
                    )}
                  </div>
                  <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4">
                    {customBayWidths.map((w, i) => (
                      <div key={i}>
                        <label className="mb-1 block text-xs text-[#ffffff]">
                          Bay {i + 1}
                          {editedBayWidths.has(i) && (
                            <span className="ml-1 text-blue-400">*</span>
                          )}
                        </label>
                        <input
                          type="number"
                          className={inputClass}
                          min={0}
                          step="0.01"
                          value={w || ''}
                          onChange={(e) =>
                            handleCustomBayWidthChange(
                              i,
                              parseFloat(e.target.value) || 0,
                            )
                          }
                        />
                      </div>
                    ))}
                  </div>

                  {/* Sum indicator */}
                  <div className="flex items-center gap-2 text-sm">
                    <span className="text-[#ffffff]">
                      Sum: {bayWidthSum.toFixed(2)}&Prime; / {openingWidth.toFixed(2)}&Prime;
                    </span>
                    {bayWidthMismatch && (
                      <span className="rounded bg-amber-900/15 px-2 py-0.5 text-xs font-medium text-yellow-400">
                        Mismatch &mdash; sum does not equal opening width
                      </span>
                    )}
                  </div>
                </div>
              )}

              {baysTall > 1 && (
                <div className="mt-3 space-y-3">
                  <div className="flex items-center gap-3">
                    <p className="text-sm font-medium text-[#ffffff]">
                      Custom Bay Heights &mdash; Bottom to Top (inches)
                    </p>
                    <button
                      type="button"
                      onClick={handleDefaultAllHeights}
                      className="rounded bg-[#1e1e2a] px-2.5 py-1 text-xs font-medium text-[#ffffff]/70 transition-colors hover:bg-[#2a2a3a] hover:text-[#ffffff]"
                    >
                      Default All
                    </button>
                    {editedBayHeights.size > 0 && editedBayHeights.size < baysTall && (
                      <button
                        type="button"
                        onClick={handleDefaultRestHeights}
                        className="rounded bg-blue-600/20 px-2.5 py-1 text-xs font-medium text-blue-400 transition-colors hover:bg-blue-600/30"
                      >
                        Default Rest
                      </button>
                    )}
                  </div>
                  <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4">
                    {[...customBayHeights].reverse().map((h, displayIdx) => {
                      const internalIdx = customBayHeights.length - 1 - displayIdx;
                      return (
                        <div key={internalIdx}>
                          <label className="mb-1 block text-xs text-[#ffffff]">
                            Bay {displayIdx + 1}
                            {displayIdx === 0 && ' (Bot)'}
                            {displayIdx === customBayHeights.length - 1 && customBayHeights.length > 1 && ' (Top)'}
                            {editedBayHeights.has(internalIdx) && (
                              <span className="ml-1 text-blue-400">*</span>
                            )}
                          </label>
                          <input
                            type="number"
                            className={inputClass}
                            min={0}
                            step="0.01"
                            value={h || ''}
                            onChange={(e) =>
                              handleCustomBayHeightChange(
                                internalIdx,
                                parseFloat(e.target.value) || 0,
                              )
                            }
                          />
                        </div>
                      );
                    })}
                  </div>

                  {/* Sum indicator */}
                  <div className="flex items-center gap-2 text-sm">
                    <span className="text-[#ffffff]">
                      Sum: {bayHeightSum.toFixed(2)}&Prime; / {openingHeight.toFixed(2)}&Prime;
                    </span>
                    {bayHeightMismatch && (
                      <span className="rounded bg-amber-900/15 px-2 py-0.5 text-xs font-medium text-yellow-400">
                        Mismatch &mdash; sum does not equal opening height
                      </span>
                    )}
                  </div>
                </div>
              )}

            </>
          )}
        </div>
      )}

      {/* ------------------------------------------------------------------ */}
      {/* Bay Diagram — directly under configuration inputs */}
      {/* ------------------------------------------------------------------ */}
      {!doorOnly && openingWidth > 0 && openingHeight > 0 && (
        <BayDiagram
          baysWide={baysWide}
          baysTall={baysTall}
          openingWidth={openingWidth}
          openingHeight={openingHeight}
          customBayWidths={baysWide > 1 ? customBayWidths : undefined}
          customBayHeights={baysTall > 1 ? customBayHeights : undefined}
          doors={doors}
        />
      )}

      {/* ------------------------------------------------------------------ */}
      {/* 5. Door Configuration */}
      {/* ------------------------------------------------------------------ */}
      <div className={cardClass}>
        <SectionHeader sectionKey="doors" title="Door Configuration" />
        {!collapsedSections.doors && (
          <>
            {doors.length === 0 && (
              <p className="text-sm text-[#ffffff]">
                No doors configured. Click &ldquo;Add Door&rdquo; to begin.
              </p>
            )}

            <div className="space-y-4">
              {doors.map((door, di) => (
                <div
                  key={di}
                  className="rounded-xl border border-[#1e1e2a] bg-[#0a0a10] p-4 space-y-3"
                >
                  <div className="flex items-center justify-between">
                    <span className="text-sm font-semibold text-[#ffffff]">
                      Door {di + 1}
                    </span>
                    <button
                      type="button"
                      onClick={() => removeDoor(di)}
                      className="rounded px-2 py-1 text-xs text-[#f87171] hover:bg-red-900/20 transition-colors duration-200"
                    >
                      Remove
                    </button>
                  </div>

                  <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
                    {/* Size */}
                    <div>
                      <label className={labelClass}>Size</label>
                      <select
                        className={selectClass}
                        value={door.size}
                        onChange={(e) =>
                          updateDoor(di, 'size', e.target.value)
                        }
                      >
                        {DOOR_SIZES.map((s) => (
                          <option key={s} value={s}>
                            {s}
                          </option>
                        ))}
                      </select>
                    </div>

                    {/* Count */}
                    <div>
                      <label className={labelClass}>Count</label>
                      <input
                        type="number"
                        className={inputClass}
                        min={1}
                        value={door.count || ''}
                        onChange={(e) =>
                          updateDoor(
                            di,
                            'count',
                            parseInt(e.target.value, 10) || 0,
                          )
                        }
                        onBlur={() => { if (door.count < 1) updateDoor(di, 'count', 1); }}
                        placeholder="1"
                      />
                    </div>

                    {/* Stile */}
                    <div>
                      <label className={labelClass}>Stile</label>
                      <select
                        className={selectClass}
                        value={door.stile}
                        onChange={(e) =>
                          updateDoor(di, 'stile', e.target.value)
                        }
                      >
                        {STILE_TYPES.map((s) => (
                          <option key={s} value={s}>
                            {s}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {/* Hardware */}
                  <div>
                    <label className={labelClass}>Hardware</label>
                    <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-3">
                      {HARDWARE_OPTIONS.map((hw) => (
                        <label
                          key={hw}
                          className="flex items-center gap-2 text-sm text-[#ffffff] cursor-pointer select-none"
                        >
                          <input
                            type="checkbox"
                            checked={door.hardware?.[hw] ?? false}
                            onChange={(e) =>
                              updateDoorHardware(di, hw, e.target.checked)
                            }
                            className="h-4 w-4 rounded border-[#ffffff] bg-[#0c0c12] text-[#3b82f6] focus:ring-[#3b82f6]/20 accent-[#3b82f6]"
                          />
                          {hw}
                        </label>
                      ))}
                    </div>
                  </div>
                </div>
              ))}
            </div>

            <button
              type="button"
              onClick={addDoor}
              className="mt-2 rounded-xl border border-dashed border-[#2a2a3a] px-4 py-2.5 text-sm text-[#ffffff] hover:border-[#3b82f6]/50 hover:text-blue-400 hover:bg-[#3b82f6]/5 transition-colors duration-200"
            >
              + Add Door
            </button>
          </>
        )}
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* 6. Installation & Field Costs (per-elevation) */}
      {/* ------------------------------------------------------------------ */}
      {!doorOnly && (
        <div className={cardClass}>
          <button
            type="button"
            onClick={() => setFieldCostExpanded(!fieldCostExpanded)}
            className="flex items-center justify-between w-full group"
          >
            <h3 className={sectionTitleClass}>
              <span className="text-amber-400/80 mr-2">$</span>
              Installation & Field Costs
            </h3>
            {fieldCostExpanded ? (
              <ChevronUp className="h-4 w-4 text-[#ffffff] group-hover:text-white transition-colors" />
            ) : (
              <ChevronDown className="h-4 w-4 text-[#ffffff] group-hover:text-white transition-colors" />
            )}
          </button>

          {fieldCostExpanded && (
            <div className="space-y-4 pt-2">
              <p className="text-xs text-[#ffffff]">
                Per-elevation field cost quantities. Rates & markups are configured in the Pricing tab.
              </p>

              {/* Installation Labor Hours */}
              <div>
                <label className={labelClass}>Installation Labor Hours</label>
                <div className="flex items-center gap-3">
                  <input
                    type="number"
                    min={0}
                    step={0.5}
                    value={installationLaborHours || ''}
                    onChange={(e) => setInstallationLaborHours(parseFloat(e.target.value) || 0)}
                    placeholder="0"
                    className={inputClass + ' max-w-[180px]'}
                  />
                  <span className="text-xs text-[#ffffff]">
                    hrs x {totalCount || 1} elev = {((installationLaborHours || 0) * (totalCount || 1)).toFixed(1)} total hrs
                  </span>
                </div>
              </div>

              {/* Sealant Joints */}
              <div>
                <label className={labelClass}>Perimeter Sealant Joints</label>
                <div className="flex items-center gap-3">
                  <input
                    type="number"
                    min={0}
                    step={1}
                    value={sealantJoints || ''}
                    onChange={(e) => setSealantJoints(parseInt(e.target.value, 10) || 0)}
                    placeholder="0"
                    className={inputClass + ' max-w-[180px]'}
                  />
                  <span className="text-xs text-[#ffffff]">
                    joints &times; {((2 * (openingWidth + openingHeight)) / 12).toFixed(1)} ft perimeter
                  </span>
                </div>
              </div>

              {/* Break Metal Selections */}
              <div>
                <label className={labelClass}>Aluminum Break Metal</label>
                <div className="flex flex-wrap gap-2">
                  {['Perimeter', 'Head', 'Sill', 'Left Jamb', 'Right Jamb', 'Both Jambs'].map((opt) => {
                    const selected = breakMetalSelections.includes(opt);
                    return (
                      <button
                        key={opt}
                        type="button"
                        onClick={() => {
                          setBreakMetalSelections((prev) => {
                            if (selected) return prev.filter((s) => s !== opt);
                            if (opt === 'Perimeter') return ['Perimeter'];
                            let next = prev.filter((s) => s !== 'Perimeter');
                            if (opt === 'Both Jambs') {
                              next = next.filter((s) => s !== 'Left Jamb' && s !== 'Right Jamb');
                            } else if (opt === 'Left Jamb' || opt === 'Right Jamb') {
                              next = next.filter((s) => s !== 'Both Jambs');
                            }
                            return [...next, opt];
                          });
                        }}
                        className={`px-3 py-1.5 rounded-lg text-xs font-medium border transition-colors duration-150 ${
                          selected
                            ? 'bg-amber-500/20 border-amber-500/50 text-amber-300'
                            : 'bg-[#0c0c12] border-[#2a2a3a] text-[#ffffff] hover:border-amber-500/30'
                        }`}
                      >
                        {opt}
                      </button>
                    );
                  })}
                </div>
                {breakMetalSelections.length > 0 && (
                  <p className="text-xs text-[#ffffff] mt-2">
                    Linear footage:{' '}
                    {(() => {
                      const wFt = openingWidth / 12;
                      const hFt = openingHeight / 12;
                      let total = 0;
                      for (const sel of breakMetalSelections) {
                        if (sel === 'Perimeter') total += 2 * (wFt + hFt);
                        else if (sel === 'Head') total += wFt;
                        else if (sel === 'Sill') total += wFt;
                        else if (sel === 'Left Jamb') total += hFt;
                        else if (sel === 'Right Jamb') total += hFt;
                        else if (sel === 'Both Jambs') total += 2 * hFt;
                      }
                      return total.toFixed(1);
                    })()}{' '}
                    ft/elev &times; {totalCount || 1} ={' '}
                    {(() => {
                      const wFt = openingWidth / 12;
                      const hFt = openingHeight / 12;
                      let total = 0;
                      for (const sel of breakMetalSelections) {
                        if (sel === 'Perimeter') total += 2 * (wFt + hFt);
                        else if (sel === 'Head') total += wFt;
                        else if (sel === 'Sill') total += wFt;
                        else if (sel === 'Left Jamb') total += hFt;
                        else if (sel === 'Right Jamb') total += hFt;
                        else if (sel === 'Both Jambs') total += 2 * hFt;
                      }
                      return (total * (totalCount || 1)).toFixed(1);
                    })()}{' '}
                    total ft
                  </p>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ------------------------------------------------------------------ */}
      {/* 7. Save / Update */}
      {/* ------------------------------------------------------------------ */}
      <div className={cardClass}>
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <h3 className={sectionTitleClass}>{isExisting ? 'Update Elevation' : 'Save Elevation'}</h3>
            {justSaved && (
              <span className="flex items-center gap-1.5 text-xs text-emerald-400 animate-fade-in">
                <CheckCircle2 className="h-4 w-4" />
                Saved successfully
              </span>
            )}
          </div>
          <button
            type="button"
            onClick={handleCalculate}
            disabled={isCalculating || !isFormValid}
            className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-[#3b82f6] to-[#2563eb] hover:brightness-110 px-6 py-2.5 text-sm font-semibold text-white disabled:opacity-40 disabled:cursor-not-allowed transition-colors duration-200"
          >
            {isCalculating ? (
              <>
                <Calculator className="h-4 w-4 animate-spin" />
                {isExisting ? 'Updating...' : 'Saving...'}
              </>
            ) : (
              <>
                <Save className="h-4 w-4" />
                {isExisting ? 'Update' : 'Save'}
              </>
            )}
          </button>
        </div>

        {!isFormValid && (
          <div className="flex items-start gap-2 rounded-lg border border-amber-500/30 bg-amber-500/10 px-4 py-3">
            <AlertTriangle className="h-4 w-4 text-amber-400 mt-0.5 shrink-0" />
            <div className="text-sm text-amber-300">
              {doorOnly ? (
                <span>Add at least one door before saving.</span>
              ) : (
                <>
                  <span className="font-medium">Required fields missing:</span>{' '}
                  {missingFields.join(', ')}
                </>
              )}
            </div>
          </div>
        )}

        {isFormValid && !results && (
          <p className="text-sm text-[#ffffff]">
            Configure the elevation above, then click &ldquo;{isExisting ? 'Update' : 'Save'}&rdquo; to price all materials and save.
          </p>
        )}

        {/* Compact cost summary after calculation */}
        {results && results.length > 0 && (
          <div className="mt-3 rounded-xl border border-[#1e1e2a] bg-[#0a0a10] p-5">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-5">
                <div>
                  <p className="text-[10px] text-[#ffffff] font-semibold uppercase tracking-wider mb-1">List Price</p>
                  <p className="text-xl font-bold font-mono text-white tabular-nums">
                    {formatCurrency(grandTotal)}
                  </p>
                </div>
                <div className="w-px h-12 bg-[#1e1e2a]" />
                <div>
                  <p className="text-[10px] text-[#ffffff] font-semibold uppercase tracking-wider mb-1">Elevation Total</p>
                  <p className="text-xl font-bold font-mono text-emerald-400 tabular-nums">
                    {formatCurrency(
                      groupedResults.reduce((s, g) => {
                        const sectionTotal = g.items.reduce((a, r) => a + (r.price ?? 0), 0);
                        return s + (g.isDiscountable ? sectionTotal * discountMultiplier : sectionTotal);
                      }, 0)
                      + ((installationLaborHours || 0) * (settings.installation_labor_rate ?? 65) * (1 + (settings.installation_labor_markup_pct ?? 0) / 100))
                      + ((sealantJoints || 0) * (settings.sealant_rate_per_ft ?? 3.5) * ((2 * (openingWidth + openingHeight)) / 12) * (1 + (settings.sealant_markup_pct ?? 0) / 100))
                    )}
                  </p>
                </div>
                <div className="w-px h-12 bg-[#1e1e2a]" />
                <div>
                  <p className="text-[10px] text-[#ffffff] font-semibold uppercase tracking-wider mb-1">Line Items</p>
                  <p className="text-xl font-bold text-[#ffffff] tabular-nums">
                    {results.filter(r => r.type !== 'Calculations').length}
                  </p>
                </div>
                {doors.length > 0 && (
                  <>
                    <div className="w-px h-12 bg-[#1e1e2a]" />
                    <div>
                      <p className="text-[10px] text-[#ffffff] font-semibold uppercase tracking-wider mb-1">Doors</p>
                      <p className="text-xl font-bold text-purple-400 tabular-nums">
                        {doors.reduce((s, d) => s + d.count, 0)}
                      </p>
                    </div>
                  </>
                )}
              </div>
              <button
                type="button"
                onClick={() => setResultsTableExpanded(!resultsTableExpanded)}
                className="flex items-center gap-2 text-xs text-[#ffffff] hover:text-[#ffffff] transition-colors duration-200"
              >
                <Table className="h-3.5 w-3.5" />
                {resultsTableExpanded ? 'Hide Details' : 'Show Details'}
                {resultsTableExpanded ? (
                  <ChevronUp className="h-3.5 w-3.5" />
                ) : (
                  <ChevronDown className="h-3.5 w-3.5" />
                )}
              </button>
            </div>

            {/* Per-category cost breakdown (matches Excel COST/ELEVATION) */}
            {groupedResults.length > 0 && (() => {
              const laborRate = settings.installation_labor_rate ?? 65;
              const laborMkp = 1 + (settings.installation_labor_markup_pct ?? 0) / 100;
              const sealRate = settings.sealant_rate_per_ft ?? 3.5;
              const sealMkp = 1 + (settings.sealant_markup_pct ?? 0) / 100;
              const perimFt = (2 * (openingWidth + openingHeight)) / 12;
              const installCost = (installationLaborHours || 0) * laborRate * laborMkp;
              const sealCost = (sealantJoints || 0) * sealRate * perimFt * sealMkp;

              const categories = groupedResults.map((g) => {
                const listTotal = g.items.reduce((a, r) => a + (r.price ?? 0), 0);
                const discTotal = g.isDiscountable ? listTotal * discountMultiplier : listTotal;
                return { label: g.label, listTotal, discTotal };
              });

              const elevTotal = categories.reduce((a, c) => a + c.discTotal, 0) + installCost + sealCost;

              return (
                <div className="mt-4 pt-3 border-t border-[#1e1e2a]">
                  <p className="text-[10px] text-[#ffffff]/60 font-semibold uppercase tracking-wider mb-2">Cost / Elevation</p>
                  <div className="grid grid-cols-2 gap-x-6 gap-y-1">
                    {categories.map((c) => (
                      <div key={c.label} className="flex items-center justify-between text-xs">
                        <span className="text-[#ffffff]/70">{c.label}</span>
                        <span className="font-mono tabular-nums text-[#ffffff]">{formatCurrency(c.discTotal)}</span>
                      </div>
                    ))}
                    {installCost > 0 && (
                      <div className="flex items-center justify-between text-xs">
                        <span className="text-[#ffffff]/70">Installation Labor</span>
                        <span className="font-mono tabular-nums text-[#ffffff]">{formatCurrency(installCost)}</span>
                      </div>
                    )}
                    {sealCost > 0 && (
                      <div className="flex items-center justify-between text-xs">
                        <span className="text-[#ffffff]/70">Perimeter Sealants</span>
                        <span className="font-mono tabular-nums text-[#ffffff]">{formatCurrency(sealCost)}</span>
                      </div>
                    )}
                  </div>
                  <div className="flex items-center justify-between mt-2 pt-2 border-t border-[#1e1e2a]">
                    <span className="text-xs font-semibold text-white">Elev Total</span>
                    <span className="text-sm font-bold font-mono tabular-nums text-emerald-400">{formatCurrency(elevTotal)}</span>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* Messages from calculation (e.g. glass area adjustments) */}
        {results && results.filter(r => r.message).length > 0 && (
          <div className="mt-2 space-y-1">
            {results.filter(r => r.message).map((r, i) => (
              <p
                key={`msg-${i}`}
                className="rounded bg-amber-900/10 px-3 py-2 text-xs text-yellow-400"
              >
                {r.description}: {r.message}
              </p>
            ))}
          </div>
        )}

        {/* ---------------------------------------------------------------- */}
        {/* Detailed results table (collapsible) */}
        {/* ---------------------------------------------------------------- */}
        {results && results.length > 0 && resultsTableExpanded && (
          <div className="mt-3 space-y-3 animate-fade-in">
            {/* Column visibility picker */}
            <div className="flex items-center justify-between">
              <p className="text-xs font-medium text-[#ffffff]">
                Material Stock List
              </p>
              <div className="relative">
                <button
                  type="button"
                  onClick={() => setShowColumnPicker(!showColumnPicker)}
                  className="flex items-center gap-1.5 rounded-md border border-[#1e1e2a] bg-[#0c0c12] px-3 py-1.5 text-xs text-[#ffffff] hover:border-[#3b82f6]/40 hover:text-[#ffffff] transition-colors duration-200"
                >
                  {showColumnPicker ? (
                    <EyeOff className="h-3 w-3" />
                  ) : (
                    <Eye className="h-3 w-3" />
                  )}
                  Columns ({visibleColumns.size}/{RESULTS_COLUMN_DEFS.filter(c => !c.perElev || totalCount > 1).length})
                </button>
                {showColumnPicker && (
                  <div className="absolute right-0 top-full z-20 mt-1 w-72 rounded-lg border border-[#1e1e2a] bg-[#111118] p-3 shadow-xl shadow-black/30">
                    <p className="mb-2 text-xs font-medium text-[#ffffff] uppercase tracking-wider">
                      Toggle Columns
                    </p>
                    <div className="space-y-1">
                      {RESULTS_COLUMN_DEFS
                        .filter((col) => !col.perElev || totalCount > 1)
                        .map((col) => (
                          <label
                            key={col.key}
                            className="flex items-center gap-2 rounded-md px-2 py-1.5 text-xs cursor-pointer hover:bg-[#16161f] transition-colors duration-150"
                          >
                            <input
                              type="checkbox"
                              checked={visibleColumns.has(col.key)}
                              onChange={() => toggleColumnVisibility(col.key)}
                              className="h-3.5 w-3.5 rounded border-[#ffffff] bg-[#0c0c12] text-[#3b82f6] focus:ring-[#3b82f6]/20 accent-[#3b82f6]"
                            />
                            <span className={visibleColumns.has(col.key) ? 'text-[#ffffff]' : 'text-[#ffffff]'}>
                              {col.label}
                            </span>
                          </label>
                        ))}
                    </div>
                    {totalCount > 1 && (
                      <p className="mt-2 text-[10px] text-[#ffffff] italic">
                        &quot;Per Elevation&quot; columns visible because count &gt; 1
                      </p>
                    )}
                  </div>
                )}
              </div>
            </div>

            {/* Grouped material tables */}
            {groupedResults.map(({ section, label, items, isDiscountable }) => {
              // Section totals
              const sectionListTotal = items.reduce((s, r) => s + (r.price ?? 0), 0);
              const sectionDiscountedTotal = isDiscountable
                ? sectionListTotal * discountMultiplier
                : sectionListTotal;

              return (
                <div key={section} className="rounded-xl border border-[#1e1e2a] bg-[#0a0a10] overflow-hidden">
                  {/* Section header */}
                  <div className="flex items-center justify-between border-b border-[#1e1e2a] bg-[#0c0c14] px-4 py-3">
                    <h4 className="text-xs font-semibold text-[#ffffff] uppercase tracking-[0.1em]">
                      {label}
                    </h4>
                    <div className="flex items-center gap-3 text-xs font-mono tabular-nums">
                      {visibleColumns.has('total_list_cost') && (
                        <span className="text-[#ffffff]">
                          List: <span className="text-white">{formatCurrency(sectionListTotal)}</span>
                        </span>
                      )}
                      {visibleColumns.has('discounted_total_list_cost') && (
                        <span className="text-[#ffffff]">
                          Disc: <span className="text-emerald-400">{formatCurrency(sectionDiscountedTotal)}</span>
                          {isDiscountable && (
                            <span className="ml-1 text-[#ffffff]">
                              ({(discountMultiplier * 100).toFixed(1)}%)
                            </span>
                          )}
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Table */}
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="border-b border-[#1e1e2a]">
                          {activeColumns.map((col) => (
                            <th
                              key={col.key}
                              className={`px-3 py-2 text-left font-medium text-[#ffffff] uppercase tracking-wider whitespace-nowrap ${
                                col.key.includes('cost') || col.key.includes('quantity') || col.key === 'quantity_per_elevation'
                                  ? 'text-right'
                                  : ''
                              }`}
                            >
                              {col.label}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {items.map((item, idx) => {
                          const listCost = item.price ?? 0;
                          const discountedCost = isDiscountable
                            ? listCost * discountMultiplier
                            : listCost;
                          const qtyRaw = item.quantity;
                          const qtyTotal = Array.isArray(qtyRaw) ? sumArray(qtyRaw) : Number(qtyRaw);

                          // Per-elevation: use single-elev data when available
                          const sKey = `${section}|${item.description}|${item.part_number}`;
                          const sData = singleElevMap.get(sKey);
                          let qtyPerElev: number;
                          let listCostPerElev: number;
                          let discountedCostPerElev: number;
                          let qtyPerElevDisplay: string;

                          if (sData && totalCount > 1) {
                            qtyPerElev = Array.isArray(sData.quantity) ? sumArray(sData.quantity) : Number(sData.quantity);
                            listCostPerElev = sData.price;
                            discountedCostPerElev = isDiscountable ? sData.price * discountMultiplier : sData.price;
                            qtyPerElevDisplay = formatQuantity(sData.quantity, section === 'gaskets' ? 'profiles' : item.type, item.unit);
                          } else if (totalCount > 1) {
                            qtyPerElev = qtyTotal / totalCount;
                            listCostPerElev = listCost / totalCount;
                            discountedCostPerElev = discountedCost / totalCount;
                            qtyPerElevDisplay = formatQuantity(
                              Array.isArray(qtyRaw)
                                ? qtyRaw.map((v) => v / totalCount)
                                : Number(qtyRaw) / totalCount,
                              section === 'gaskets' ? 'profiles' : item.type,
                              item.unit,
                            );
                          } else {
                            qtyPerElev = qtyTotal;
                            listCostPerElev = listCost;
                            discountedCostPerElev = discountedCost;
                            qtyPerElevDisplay = '';
                          }

                          // Format quantity based on type
                          const qtyDisplay = formatQuantity(qtyRaw, section === 'gaskets' ? 'profiles' : item.type, item.unit);

                          // Cell value lookup
                          const cellValue: Record<string, React.ReactNode> = {
                            description: (
                              <span className="text-[#ffffff]" title={item.description}>
                                {item.description}
                              </span>
                            ),
                            part_number: (
                              <span className="font-mono text-[#ffffff]">{item.part_number}</span>
                            ),
                            total_quantity_required: (
                              <span className="font-mono text-[#ffffff] tabular-nums">{qtyDisplay}</span>
                            ),
                            quantity_per_elevation: (
                              <span className="font-mono text-[#ffffff] tabular-nums">{qtyPerElevDisplay}</span>
                            ),
                            total_list_cost: (
                              <span className="font-mono text-white tabular-nums">{formatCurrency(listCost)}</span>
                            ),
                            total_list_cost_per_elevation: (
                              <span className="font-mono text-[#ffffff] tabular-nums">{formatCurrency(listCostPerElev)}</span>
                            ),
                            discounted_total_list_cost: (
                              <span className={`font-mono tabular-nums ${isDiscountable ? 'text-emerald-400' : 'text-white'}`}>
                                {formatCurrency(discountedCost)}
                              </span>
                            ),
                            discounted_total_list_cost_per_elevation: (
                              <span className={`font-mono tabular-nums ${isDiscountable ? 'text-emerald-400/70' : 'text-[#ffffff]'}`}>
                                {formatCurrency(discountedCostPerElev)}
                              </span>
                            ),
                          };

                          return (
                            <tr
                              key={`${section}-${idx}`}
                              className="border-b border-[#1e1e2a]/50 hover:bg-[#111118] transition-colors duration-100"
                            >
                              {activeColumns.map((col) => (
                                <td
                                  key={col.key}
                                  className={`px-3 py-2 whitespace-nowrap ${
                                    col.key.includes('cost') || col.key.includes('quantity') || col.key === 'quantity_per_elevation'
                                      ? 'text-right'
                                      : ''
                                  }`}
                                >
                                  {cellValue[col.key] ?? '—'}
                                </td>
                              ))}
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })}

            {/* Grand total row */}
            {groupedResults.length > 0 && (
              <div className="rounded-xl border border-[#1e1e2a] bg-[#0c0c14] px-5 py-4">
                <div className="flex items-center justify-between">
                  <span className="text-sm font-semibold text-white">
                    Elevation Total
                  </span>
                  <div className="flex items-center gap-4 text-sm font-mono tabular-nums">
                    {visibleColumns.has('total_list_cost') && (
                      <span className="text-[#ffffff]">
                        List: <span className="font-bold text-white">{formatCurrency(grandTotal)}</span>
                      </span>
                    )}
                    {visibleColumns.has('discounted_total_list_cost') && (
                      <span className="text-[#ffffff]">
                        Discounted:{' '}
                        <span className="font-bold text-emerald-400">
                          {formatCurrency(
                            groupedResults.reduce((s, g) => {
                              const sectionTotal = g.items.reduce((a, r) => a + (r.price ?? 0), 0);
                              return s + (g.isDiscountable ? sectionTotal * discountMultiplier : sectionTotal);
                            }, 0),
                          )}
                        </span>
                      </span>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
