'use client';

import { useState, useMemo, useCallback, useEffect } from 'react';
import {
  Save,
  RefreshCw,
  ChevronDown,
  ChevronUp,
  DollarSign,
  TrendingUp,
  BarChart3,
  AlertTriangle,
  Info,
  Loader2,
  Package,
} from 'lucide-react';
import type {
  ElevationData,
  ExtraMaterial,
  ProjectSettings,
  WasteMaterialBreakdown,
  WasteSuggestion,
} from '@/types';
import { getUnitPriceByPart } from '@/lib/pricing';
import { partsData } from '@/data/parts-data';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

interface SummaryTabProps {
  elevations: Record<string, ElevationData>;
  materials: Record<string, ExtraMaterial>;
  settings: ProjectSettings;
  onSettingsUpdate: (newSettings: ProjectSettings) => Promise<void>;
}

// ---------------------------------------------------------------------------
// Shared styling (matching ElevationEditor)
// ---------------------------------------------------------------------------

const inputClass =
  'bg-[#1c1c21] border border-[#27272a] text-white rounded-lg px-3 py-2 w-full focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500 transition-colors text-sm';
const labelClass = 'block text-sm font-medium text-zinc-400 mb-1';
const cardClass = 'bg-[#18181b] border border-[#27272a] rounded-xl p-5 space-y-4';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function formatCurrency(value: number): string {
  return '$' + value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseKey(key: string): { partNumber: string; finish?: string } {
  const lastDashIdx = key.lastIndexOf('-');
  if (lastDashIdx > 0) {
    const possibleFinish = key.substring(lastDashIdx + 1);
    if (['clear', 'black', 'paint', 'bronze', 'mill'].includes(possibleFinish)) {
      return { partNumber: key.substring(0, lastDashIdx), finish: possibleFinish };
    }
  }
  return { partNumber: key };
}

function getWasteColor(pct: number): string {
  if (pct < 10) return 'text-emerald-400';
  if (pct < 20) return 'text-yellow-400';
  return 'text-red-400';
}

function getWasteBarColor(pct: number): string {
  if (pct < 10) return 'bg-emerald-500';
  if (pct < 20) return 'bg-yellow-500';
  return 'bg-red-500';
}

function getWasteProgressColor(pct: number): string {
  if (pct < 10) return 'bg-emerald-500';
  if (pct < 20) return 'bg-yellow-500';
  if (pct < 30) return 'bg-orange-500';
  return 'bg-red-500';
}

function getPriorityStyles(priority: 'high' | 'medium' | 'low'): string {
  switch (priority) {
    case 'high': return 'border-red-500/40 bg-red-500/5';
    case 'medium': return 'border-yellow-500/40 bg-yellow-500/5';
    case 'low': return 'border-emerald-500/40 bg-emerald-500/5';
  }
}

function getPriorityBadge(priority: 'high' | 'medium' | 'low'): string {
  switch (priority) {
    case 'high': return 'bg-red-500/20 text-red-400';
    case 'medium': return 'bg-yellow-500/20 text-yellow-400';
    case 'low': return 'bg-emerald-500/20 text-emerald-400';
  }
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export default function SummaryTab({
  elevations,
  materials,
  settings,
  onSettingsUpdate,
}: SummaryTabProps) {
  // ---- Section collapse state ----
  const [collapsedSections, setCollapsedSections] = useState<Record<string, boolean>>({});
  const toggleSection = (key: string) =>
    setCollapsedSections((prev) => ({ ...prev, [key]: !prev[key] }));

  // ---- Save indicators ----
  const [savingAdditional, setSavingAdditional] = useState(false);
  const [savingMarkups, setSavingMarkups] = useState(false);
  const [savingDisplay, setSavingDisplay] = useState(false);

  // ---- Additional Cost Settings state ----
  const [overheadMaterialsPct, setOverheadMaterialsPct] = useState(settings.overhead_materials_pct ?? 0);
  const [overheadLaborPct, setOverheadLaborPct] = useState(settings.overhead_labor_pct ?? 0);
  const [adminManagementPct, setAdminManagementPct] = useState(settings.admin_management_pct ?? 0);
  const [engineeringPct, setEngineeringPct] = useState(settings.engineering_pct ?? 0);
  const [packagingMaterialsPct, setPackagingMaterialsPct] = useState(settings.packaging_materials_pct ?? 0);
  const [shippingTransportPct, setShippingTransportPct] = useState(settings.shipping_transport_pct ?? 0);
  const [commissionsPct, setCommissionsPct] = useState(settings.commissions_pct ?? 0);

  // ---- Markup Settings state ----
  const [profitOnMaterialPct, setProfitOnMaterialPct] = useState(settings.profit_on_material_pct ?? 0);
  const [profitOnWastePct, setProfitOnWastePct] = useState(settings.profit_on_waste_pct ?? 0);
  const [profitOnGlassPct, setProfitOnGlassPct] = useState(settings.profit_on_glass_pct ?? 0);
  const [profitOnWagesPct, setProfitOnWagesPct] = useState(settings.profit_on_wages_pct ?? 0);
  const [planningTechnicalPct, setPlanningTechnicalPct] = useState(settings.planning_technical_pct ?? 0);
  const [commissionPct, setCommissionPct] = useState(settings.commission_pct ?? 0);

  // ---- Elevation Summary Display state ----
  const [showElevationNames, setShowElevationNames] = useState(settings.show_elevation_names ?? false);
  const [showElevationQuantity, setShowElevationQuantity] = useState(settings.show_elevation_quantity ?? false);
  const [showElevationDimensions, setShowElevationDimensions] = useState(settings.show_elevation_dimensions ?? false);
  const [showElevationSqft, setShowElevationSqft] = useState(settings.show_elevation_sqft ?? false);
  const [showElevationPerimeter, setShowElevationPerimeter] = useState(settings.show_elevation_perimeter ?? false);

  // Sync from props when settings change externally
  useEffect(() => {
    setOverheadMaterialsPct(settings.overhead_materials_pct ?? 0);
    setOverheadLaborPct(settings.overhead_labor_pct ?? 0);
    setAdminManagementPct(settings.admin_management_pct ?? 0);
    setEngineeringPct(settings.engineering_pct ?? 0);
    setPackagingMaterialsPct(settings.packaging_materials_pct ?? 0);
    setShippingTransportPct(settings.shipping_transport_pct ?? 0);
    setCommissionsPct(settings.commissions_pct ?? 0);
    setProfitOnMaterialPct(settings.profit_on_material_pct ?? 0);
    setProfitOnWastePct(settings.profit_on_waste_pct ?? 0);
    setProfitOnGlassPct(settings.profit_on_glass_pct ?? 0);
    setProfitOnWagesPct(settings.profit_on_wages_pct ?? 0);
    setPlanningTechnicalPct(settings.planning_technical_pct ?? 0);
    setCommissionPct(settings.commission_pct ?? 0);
    setShowElevationNames(settings.show_elevation_names ?? false);
    setShowElevationQuantity(settings.show_elevation_quantity ?? false);
    setShowElevationDimensions(settings.show_elevation_dimensions ?? false);
    setShowElevationSqft(settings.show_elevation_sqft ?? false);
    setShowElevationPerimeter(settings.show_elevation_perimeter ?? false);
  }, [settings]);

  // ---- Save Additional Costs ----
  const handleSaveAdditionalCosts = useCallback(async () => {
    setSavingAdditional(true);
    try {
      await onSettingsUpdate({
        ...settings,
        overhead_materials_pct: overheadMaterialsPct,
        overhead_labor_pct: overheadLaborPct,
        admin_management_pct: adminManagementPct,
        engineering_pct: engineeringPct,
        packaging_materials_pct: packagingMaterialsPct,
        shipping_transport_pct: shippingTransportPct,
        commissions_pct: commissionsPct,
      });
    } finally {
      setSavingAdditional(false);
    }
  }, [
    settings, onSettingsUpdate,
    overheadMaterialsPct, overheadLaborPct, adminManagementPct,
    engineeringPct, packagingMaterialsPct, shippingTransportPct, commissionsPct,
  ]);

  // ---- Save Markups ----
  const handleSaveMarkups = useCallback(async () => {
    setSavingMarkups(true);
    try {
      await onSettingsUpdate({
        ...settings,
        profit_on_material_pct: profitOnMaterialPct,
        profit_on_waste_pct: profitOnWastePct,
        profit_on_glass_pct: profitOnGlassPct,
        profit_on_wages_pct: profitOnWagesPct,
        planning_technical_pct: planningTechnicalPct,
        commission_pct: commissionPct,
      });
    } finally {
      setSavingMarkups(false);
    }
  }, [
    settings, onSettingsUpdate,
    profitOnMaterialPct, profitOnWastePct, profitOnGlassPct,
    profitOnWagesPct, planningTechnicalPct, commissionPct,
  ]);

  // ---- Save Elevation Summary Display ----
  const handleSaveDisplay = useCallback(async () => {
    setSavingDisplay(true);
    try {
      await onSettingsUpdate({
        ...settings,
        show_elevation_names: showElevationNames,
        show_elevation_quantity: showElevationQuantity,
        show_elevation_dimensions: showElevationDimensions,
        show_elevation_sqft: showElevationSqft,
        show_elevation_perimeter: showElevationPerimeter,
      });
    } finally {
      setSavingDisplay(false);
    }
  }, [
    settings, onSettingsUpdate,
    showElevationNames, showElevationQuantity, showElevationDimensions,
    showElevationSqft, showElevationPerimeter,
  ]);

  // ---- Waste Calculator ----
  const wasteAnalysis = useMemo(() => {
    const usageMap = new Map<string, {
      partNumber: string;
      finish: string;
      description: string;
      totalRequested: number;
      totalPurchased: number;
      totalUsedFromLeftover: number;
      totalCostIncurred: number;
      unit: string;
    }>();

    for (const elev of Object.values(elevations)) {
      if (!elev.material_impacts) continue;
      for (const impact of elev.material_impacts) {
        const pn = impact.part_number;
        if (!pn || pn === 'N/A - Manual') continue;
        const finish = impact.finish || elev.finish || 'clear';
        const key = `${pn}-${finish.toLowerCase()}`;
        const existing = usageMap.get(key);
        const requestedQty = Array.isArray(impact.requested_qty)
          ? impact.requested_qty.reduce((s, v) => s + Number(v), 0)
          : Number(impact.requested_qty);
        const unit = impact.type_processed_as === 'profile' ? 'ft' : 'pcs';
        if (existing) {
          existing.totalRequested += requestedQty;
          existing.totalPurchased += impact.purchased_qty_or_length;
          existing.totalUsedFromLeftover += impact.used_from_leftover_qty_or_length;
          existing.totalCostIncurred += impact.cost_incurred;
        } else {
          usageMap.set(key, {
            partNumber: pn, finish, description: impact.description || pn,
            totalRequested: requestedQty,
            totalPurchased: impact.purchased_qty_or_length,
            totalUsedFromLeftover: impact.used_from_leftover_qty_or_length,
            totalCostIncurred: impact.cost_incurred,
            unit,
          });
        }
      }
    }

    const breakdown: WasteMaterialBreakdown[] = [];
    let totalWasteCost = 0;
    let totalMaterialCost = 0;

    for (const [key, usage] of usageMap) {
      const { partNumber, finish } = parseKey(key);
      const [unitPrice] = getUnitPriceByPart(partNumber, finish);
      const matEntry = materials[key];
      const leftoverPieces = matEntry?.length_pieces ?? [];
      const leftoverQty = matEntry?.quantity ?? 0;

      let wasteQty: number;
      let wasteDisplay: string;
      if (usage.unit === 'ft') {
        wasteQty = leftoverPieces.reduce((s, l) => s + l, 0);
        wasteDisplay = wasteQty.toFixed(2) + ' ft';
      } else {
        wasteQty = leftoverQty;
        wasteDisplay = wasteQty.toFixed(0) + ' pcs';
      }

      const wasteCost = (unitPrice ?? 0) * wasteQty;
      const totalUsed = usage.totalRequested;
      const totalAcquired = totalUsed + wasteQty;
      const wastePercentage = totalAcquired > 0
        ? Math.min((wasteQty / totalAcquired) * 100, 100.0)
        : 0;

      breakdown.push({
        part_number: partNumber,
        description: usage.description || partNumber,
        finish: usage.finish,
        total_quantity: totalUsed,
        waste_quantity: wasteQty,
        waste_quantity_display: wasteDisplay,
        waste_percentage: wastePercentage,
        waste_cost: wasteCost,
        material_cost: usage.totalCostIncurred,
        unit: usage.unit,
        individual_pieces: leftoverPieces,
      });
      totalWasteCost += wasteCost;
      totalMaterialCost += usage.totalCostIncurred;
    }

    breakdown.sort((a, b) => b.waste_cost - a.waste_cost);

    const overallWastePercentage = totalMaterialCost > 0
      ? (totalWasteCost / totalMaterialCost) * 100
      : 0;

    // Generate suggestions
    const suggestions: WasteSuggestion[] = [];
    const highWasteItems = breakdown.filter(m => m.waste_percentage > 30 && m.waste_cost > 50);
    for (const item of highWasteItems.slice(0, 3)) {
      suggestions.push({
        priority: 'high', category: 'High Waste',
        message: `${item.part_number} (${item.finish}) has ${item.waste_percentage.toFixed(1)}% waste. Consider adjusting bay dimensions.`,
        estimated_savings: item.waste_cost * 0.5,
      });
    }
    const medWaste = breakdown.filter(m => m.waste_percentage > 15 && m.waste_percentage <= 30 && m.waste_cost > 25);
    for (const item of medWaste.slice(0, 2)) {
      suggestions.push({
        priority: 'medium', category: 'Moderate Waste',
        message: `${item.part_number} has ${item.waste_percentage.toFixed(1)}% waste (${formatCurrency(item.waste_cost)}). Leftover pieces may be usable in future elevations.`,
        estimated_savings: item.waste_cost * 0.3,
      });
    }
    const reusable = breakdown.filter(m => m.individual_pieces.length > 0 && m.waste_percentage <= 15);
    if (reusable.length > 0) {
      suggestions.push({
        priority: 'low', category: 'Reuse Opportunity',
        message: `${reusable.length} material(s) have small leftover pieces that could be reused across elevations.`,
        estimated_savings: null,
      });
    }
    if (overallWastePercentage > 20) {
      suggestions.push({
        priority: 'high', category: 'Overall Waste',
        message: `Overall waste is ${overallWastePercentage.toFixed(1)}% (${formatCurrency(totalWasteCost)}). Review opening dimensions and bay layouts.`,
        estimated_savings: totalWasteCost * 0.3,
      });
    }

    return { breakdown, totalWasteCost, totalMaterialCost, overallWastePercentage, suggestions };
  }, [elevations, materials]);

  // ---- Elevation Summary ----
  const elevationSummary = useMemo(() => {
    return Object.entries(elevations).map(([name, elev]) => {
      const wInches = elev.opening_width_inches || 0;
      const hInches = elev.opening_height_inches || 0;
      const sqft = (wInches * hInches) / 144;
      const perimeterFt = (2 * (wInches + hInches)) / 12;
      return {
        name,
        quantity: elev.total_count || 1,
        width: wInches,
        height: hInches,
        sqft: sqft * (elev.total_count || 1),
        perimeter: perimeterFt * (elev.total_count || 1),
      };
    });
  }, [elevations]);

  // ---- Stock / Leftover Inventory (cross-elevation) ----
  const stockInventory = useMemo(() => {
    const items: Array<{
      partNumber: string;
      finish: string;
      description: string;
      type: 'profile' | 'accessory';
      pieces: number[];           // individual piece lengths (for profiles)
      totalLength: number;        // sum of pieces
      quantity: number;           // leftover quantity (for accessories)
      unitPrice: number;
      estimatedValue: number;
      pieceSummary: string;       // grouped summary e.g. "8.25ft x2, 3.50ft x1"
    }> = [];

    let totalPieces = 0;
    let totalValue = 0;

    for (const [key, mat] of Object.entries(materials)) {
      const hasLengthPieces = mat.length_pieces && mat.length_pieces.length > 0;
      const hasQuantity = mat.quantity > 0;
      if (!hasLengthPieces && !hasQuantity) continue;

      const { partNumber, finish } = parseKey(key);
      const [unitPrice] = getUnitPriceByPart(partNumber, finish ?? 'clear');

      // Try to find description from elevation material impacts
      let description = partNumber;
      for (const elev of Object.values(elevations)) {
        if (!elev.material_impacts) continue;
        const found = elev.material_impacts.find(
          (m) => m.part_number === partNumber,
        );
        if (found?.description) {
          description = found.description;
          break;
        }
      }

      if (hasLengthPieces) {
        // Profile-type leftover (length pieces)
        const pieces = mat.length_pieces.filter((l) => l > 0).sort((a, b) => b - a);
        if (pieces.length === 0) continue;

        const totalLength = pieces.reduce((s, l) => s + l, 0);
        const value = (unitPrice ?? 0) * totalLength;

        // Group identical lengths: e.g., "8.25ft x2, 3.50ft x1"
        const grouped = new Map<string, number>();
        for (const p of pieces) {
          const rounded = p.toFixed(2);
          grouped.set(rounded, (grouped.get(rounded) ?? 0) + 1);
        }
        const pieceSummary = Array.from(grouped.entries())
          .map(([len, count]) => count > 1 ? `${len}ft x${count}` : `${len}ft`)
          .join(', ');

        items.push({
          partNumber,
          finish: finish ?? 'clear',
          description,
          type: 'profile',
          pieces,
          totalLength,
          quantity: pieces.length,
          unitPrice: unitPrice ?? 0,
          estimatedValue: value,
          pieceSummary,
        });

        totalPieces += pieces.length;
        totalValue += value;
      } else if (hasQuantity) {
        // Accessory-type leftover (quantity)
        const value = (unitPrice ?? 0) * mat.quantity;
        items.push({
          partNumber,
          finish: finish ?? 'clear',
          description,
          type: 'accessory',
          pieces: [],
          totalLength: 0,
          quantity: mat.quantity,
          unitPrice: unitPrice ?? 0,
          estimatedValue: value,
          pieceSummary: `${mat.quantity} pcs`,
        });
        totalPieces += mat.quantity;
        totalValue += value;
      }
    }

    // Sort by estimated value descending
    items.sort((a, b) => b.estimatedValue - a.estimatedValue);

    return { items, totalPieces, totalValue };
  }, [materials, elevations]);

  // ---- Section Header helper ----
  const SectionHeader = ({ sectionKey, title, icon, action }: {
    sectionKey: string; title: string; icon?: React.ReactNode;
    action?: React.ReactNode;
  }) => (
    <div className="flex items-center justify-between">
      <button type="button" className="flex items-center gap-2 flex-1" onClick={() => toggleSection(sectionKey)}>
        {icon}
        <h3 className="text-lg font-semibold text-white">{title}</h3>
        {collapsedSections[sectionKey] ? (
          <ChevronDown className="h-4 w-4 text-zinc-400 ml-auto" />
        ) : (
          <ChevronUp className="h-4 w-4 text-zinc-400 ml-auto" />
        )}
      </button>
      {action}
    </div>
  );

  // ---- Percentage input helper ----
  const PctInput = ({ label, value, onChange }: {
    label: string; value: number; onChange: (v: number) => void;
  }) => (
    <div>
      <label className={labelClass}>{label}</label>
      <div className="relative">
        <input
          type="number"
          className={inputClass + ' pr-8'}
          min={0}
          max={100}
          step="0.1"
          value={value || ''}
          onChange={(e) => onChange(parseFloat(e.target.value) || 0)}
          placeholder="0"
        />
        <span className="absolute right-3 top-1/2 -translate-y-1/2 text-zinc-500 text-sm">%</span>
      </div>
    </div>
  );

  return (
    <div className="mx-auto max-w-5xl space-y-6 p-6 pb-12">
      {/* ------------------------------------------------------------------ */}
      {/* 1. Additional Cost Settings */}
      {/* ------------------------------------------------------------------ */}
      <div className={cardClass}>
        <SectionHeader
          sectionKey="additional"
          title="Additional Cost Settings"
          icon={<DollarSign className="w-5 h-5 text-blue-500" />}
          action={
            <button
              onClick={handleSaveAdditionalCosts}
              disabled={savingAdditional}
              className="flex items-center gap-2 rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-40 transition-colors"
            >
              {savingAdditional ? <Loader2 className="w-4 h-4 animate-spin" /> : <Save className="w-4 h-4" />}
              Save
            </button>
          }
        />
        {!collapsedSections.additional && (
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
            <PctInput label="Overhead Materials %" value={overheadMaterialsPct} onChange={setOverheadMaterialsPct} />
            <PctInput label="Overhead Labor %" value={overheadLaborPct} onChange={setOverheadLaborPct} />
            <PctInput label="Admin and Management %" value={adminManagementPct} onChange={setAdminManagementPct} />
            <PctInput label="Engineering %" value={engineeringPct} onChange={setEngineeringPct} />
            <PctInput label="Packaging Materials %" value={packagingMaterialsPct} onChange={setPackagingMaterialsPct} />
            <PctInput label="Shipping and Transport %" value={shippingTransportPct} onChange={setShippingTransportPct} />
            <PctInput label="Commissions %" value={commissionsPct} onChange={setCommissionsPct} />
          </div>
        )}
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* 2. Markup Settings */}
      {/* ------------------------------------------------------------------ */}
      <div className={cardClass}>
        <SectionHeader
          sectionKey="markups"
          title="Markup Settings"
          icon={<TrendingUp className="w-5 h-5 text-emerald-500" />}
          action={
            <button
              onClick={handleSaveMarkups}
              disabled={savingMarkups}
              className="flex items-center gap-2 rounded-lg bg-emerald-600 px-4 py-2 text-sm font-medium text-white hover:bg-emerald-700 disabled:opacity-40 transition-colors"
            >
              {savingMarkups ? <Loader2 className="w-4 h-4 animate-spin" /> : <Save className="w-4 h-4" />}
              Save
            </button>
          }
        />
        {!collapsedSections.markups && (
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
            <PctInput label="Profit on Material %" value={profitOnMaterialPct} onChange={setProfitOnMaterialPct} />
            <PctInput label="Profit on Waste %" value={profitOnWastePct} onChange={setProfitOnWastePct} />
            <PctInput label="Profit on Glass Purchase %" value={profitOnGlassPct} onChange={setProfitOnGlassPct} />
            <PctInput label="Profit on Wages %" value={profitOnWagesPct} onChange={setProfitOnWagesPct} />
            <PctInput label="Planning / Technical Office %" value={planningTechnicalPct} onChange={setPlanningTechnicalPct} />
            <PctInput label="Commission %" value={commissionPct} onChange={setCommissionPct} />
          </div>
        )}
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* 3. Elevation Summary Display */}
      {/* ------------------------------------------------------------------ */}
      <div className={cardClass}>
        <SectionHeader
          sectionKey="elevDisplay"
          title="Elevation Summary Display"
          icon={<Info className="w-5 h-5 text-cyan-500" />}
          action={
            <button
              onClick={handleSaveDisplay}
              disabled={savingDisplay}
              className="flex items-center gap-2 rounded-lg bg-cyan-600 px-4 py-2 text-sm font-medium text-white hover:bg-cyan-700 disabled:opacity-40 transition-colors"
            >
              {savingDisplay ? <Loader2 className="w-4 h-4 animate-spin" /> : <Save className="w-4 h-4" />}
              Save
            </button>
          }
        />
        {!collapsedSections.elevDisplay && (
          <>
            <p className="text-xs text-zinc-500">
              Select which columns to display in the elevation summary table in exported reports.
            </p>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {[
                { label: 'Elevation Names', checked: showElevationNames, setter: setShowElevationNames },
                { label: 'Quantity', checked: showElevationQuantity, setter: setShowElevationQuantity },
                { label: 'Dimensions', checked: showElevationDimensions, setter: setShowElevationDimensions },
                { label: 'SQFT Total', checked: showElevationSqft, setter: setShowElevationSqft },
                { label: 'Perimeter FT Total', checked: showElevationPerimeter, setter: setShowElevationPerimeter },
              ].map(({ label, checked, setter }) => (
                <label key={label} className="flex items-center gap-2.5 text-sm text-zinc-300 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={checked}
                    onChange={(e) => setter(e.target.checked)}
                    className="h-4 w-4 rounded border-zinc-600 bg-[#1c1c21] text-blue-500 focus:ring-blue-500/50 accent-blue-500"
                  />
                  {label}
                </label>
              ))}
            </div>

            {/* Elevation Summary Table Preview */}
            {elevationSummary.length > 0 && (
              <div className="mt-4 overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-[#27272a] text-xs uppercase tracking-wider text-zinc-500">
                      {showElevationNames && <th className="pb-2 pr-4 text-left">Elevation</th>}
                      {showElevationQuantity && <th className="pb-2 pr-4 text-right">Qty</th>}
                      {showElevationDimensions && <th className="pb-2 pr-4 text-right">Dimensions</th>}
                      {showElevationSqft && <th className="pb-2 pr-4 text-right">SQFT</th>}
                      {showElevationPerimeter && <th className="pb-2 text-right">Perimeter FT</th>}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#27272a]/50">
                    {elevationSummary.map((elev) => (
                      <tr key={elev.name} className="text-zinc-300">
                        {showElevationNames && <td className="py-2 pr-4">{elev.name}</td>}
                        {showElevationQuantity && <td className="py-2 pr-4 text-right font-mono">{elev.quantity}</td>}
                        {showElevationDimensions && (
                          <td className="py-2 pr-4 text-right font-mono text-xs">
                            {elev.width.toFixed(1)}&quot; x {elev.height.toFixed(1)}&quot;
                          </td>
                        )}
                        {showElevationSqft && <td className="py-2 pr-4 text-right font-mono">{elev.sqft.toFixed(2)}</td>}
                        {showElevationPerimeter && <td className="py-2 text-right font-mono">{elev.perimeter.toFixed(2)}</td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* 4. Waste Calculator */}
      {/* ------------------------------------------------------------------ */}
      <div className={cardClass}>
        <SectionHeader
          sectionKey="waste"
          title="Waste Calculator"
          icon={<BarChart3 className="w-5 h-5 text-orange-500" />}
        />
        {!collapsedSections.waste && (
          <>
            {wasteAnalysis.breakdown.length === 0 ? (
              <div className="text-center py-8">
                <BarChart3 className="w-8 h-8 text-zinc-600 mx-auto mb-2" />
                <p className="text-sm text-zinc-500">
                  No material data available. Calculate elevations first.
                </p>
              </div>
            ) : (
              <>
                {/* Overall metrics */}
                <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
                  {/* Overall waste % with progress bar */}
                  <div className="bg-[#09090b] border border-[#27272a] rounded-lg p-4">
                    <p className="text-xs text-zinc-500 mb-1 font-medium">Overall Waste</p>
                    <p className={`text-2xl font-bold font-mono tabular-nums ${getWasteColor(wasteAnalysis.overallWastePercentage)}`}>
                      {wasteAnalysis.overallWastePercentage.toFixed(2)}%
                    </p>
                    <div className="mt-2 w-full h-2 bg-[#27272a] rounded-full overflow-hidden">
                      <div
                        className={`h-full rounded-full transition-all ${getWasteProgressColor(wasteAnalysis.overallWastePercentage)}`}
                        style={{ width: `${Math.min(wasteAnalysis.overallWastePercentage, 100)}%` }}
                      />
                    </div>
                  </div>

                  <div className="bg-[#09090b] border border-[#27272a] rounded-lg p-4">
                    <p className="text-xs text-zinc-500 mb-1 font-medium">Total Waste Cost</p>
                    <p className="text-2xl font-bold font-mono text-yellow-400 tabular-nums">
                      {formatCurrency(wasteAnalysis.totalWasteCost)}
                    </p>
                  </div>

                  <div className="bg-[#09090b] border border-[#27272a] rounded-lg p-4">
                    <p className="text-xs text-zinc-500 mb-1 font-medium">Total Material Cost</p>
                    <p className="text-2xl font-bold font-mono text-white tabular-nums">
                      {formatCurrency(wasteAnalysis.totalMaterialCost)}
                    </p>
                  </div>
                </div>

                {/* Breakdown table */}
                <div className="overflow-x-auto border border-[#27272a] rounded-lg">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-[#27272a] bg-[#111113]">
                        <th className="text-left px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider">Material</th>
                        <th className="text-right px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider w-28">Waste %</th>
                        <th className="text-right px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider">Waste Cost</th>
                        <th className="text-right px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider">Waste Qty</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#27272a]/50">
                      {wasteAnalysis.breakdown.map((mat, i) => (
                        <tr key={i} className="hover:bg-[#1c1c21] transition-colors">
                          <td className="px-4 py-2.5">
                            <div className="text-white text-xs">{mat.description}</div>
                            <div className="text-zinc-500 text-[10px] font-mono">{mat.part_number} ({mat.finish})</div>
                          </td>
                          <td className="px-4 py-2.5 text-right">
                            <div className="flex items-center justify-end gap-2">
                              <div className="w-12 h-1.5 bg-[#27272a] rounded-full overflow-hidden">
                                <div
                                  className={`h-full rounded-full ${getWasteBarColor(mat.waste_percentage)}`}
                                  style={{ width: `${Math.min(mat.waste_percentage, 100)}%` }}
                                />
                              </div>
                              <span className={`font-mono text-xs tabular-nums ${getWasteColor(mat.waste_percentage)}`}>
                                {mat.waste_percentage.toFixed(1)}%
                              </span>
                            </div>
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono text-xs text-yellow-400 tabular-nums">
                            {formatCurrency(mat.waste_cost)}
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono text-xs text-zinc-300 tabular-nums">
                            {mat.waste_quantity_display}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Suggestions */}
                {wasteAnalysis.suggestions.length > 0 && (
                  <div className="space-y-2 mt-2">
                    <div className="flex items-center gap-2">
                      <AlertTriangle className="w-4 h-4 text-orange-500" />
                      <h4 className="text-sm font-semibold text-white">Optimization Suggestions</h4>
                    </div>
                    {wasteAnalysis.suggestions.map((suggestion, i) => (
                      <div key={i} className={`border rounded-lg p-3 ${getPriorityStyles(suggestion.priority)}`}>
                        <div className="flex items-start justify-between gap-3">
                          <div className="flex-1">
                            <div className="flex items-center gap-2 mb-0.5">
                              <span className={`text-[10px] font-semibold uppercase px-1.5 py-0.5 rounded ${getPriorityBadge(suggestion.priority)}`}>
                                {suggestion.priority}
                              </span>
                              <span className="text-xs text-zinc-500">{suggestion.category}</span>
                            </div>
                            <p className="text-xs text-zinc-400 leading-relaxed">{suggestion.message}</p>
                          </div>
                          {suggestion.estimated_savings != null && (
                            <div className="text-right shrink-0">
                              <p className="text-[10px] text-zinc-500 uppercase font-medium">Est. Savings</p>
                              <p className="text-sm text-emerald-400 font-mono font-semibold tabular-nums">
                                {formatCurrency(suggestion.estimated_savings)}
                              </p>
                            </div>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}
          </>
        )}
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* 5. Stock / Leftover Inventory */}
      {/* ------------------------------------------------------------------ */}
      <div className={cardClass}>
        <SectionHeader
          sectionKey="stock"
          title="Stock / Leftover Inventory"
          icon={<Package className="w-5 h-5 text-purple-500" />}
        />
        {!collapsedSections.stock && (
          <>
            {stockInventory.items.length === 0 ? (
              <div className="text-center py-8">
                <Package className="w-8 h-8 text-zinc-600 mx-auto mb-2" />
                <p className="text-sm text-zinc-500">
                  No leftover inventory. Calculate elevations to generate material leftovers.
                </p>
              </div>
            ) : (
              <>
                {/* Summary stats */}
                <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
                  <div className="bg-[#09090b] border border-[#27272a] rounded-lg p-4">
                    <p className="text-xs text-zinc-500 mb-1 font-medium">Total Materials</p>
                    <p className="text-2xl font-bold font-mono text-purple-400 tabular-nums">
                      {stockInventory.items.length}
                    </p>
                    <p className="text-xs text-zinc-500 mt-0.5">unique part numbers</p>
                  </div>

                  <div className="bg-[#09090b] border border-[#27272a] rounded-lg p-4">
                    <p className="text-xs text-zinc-500 mb-1 font-medium">Total Pieces</p>
                    <p className="text-2xl font-bold font-mono text-blue-400 tabular-nums">
                      {stockInventory.totalPieces}
                    </p>
                    <p className="text-xs text-zinc-500 mt-0.5">leftover pieces/units</p>
                  </div>

                  <div className="bg-[#09090b] border border-[#27272a] rounded-lg p-4">
                    <p className="text-xs text-zinc-500 mb-1 font-medium">Estimated Value</p>
                    <p className="text-2xl font-bold font-mono text-emerald-400 tabular-nums">
                      {formatCurrency(stockInventory.totalValue)}
                    </p>
                    <p className="text-xs text-zinc-500 mt-0.5">reusable material value</p>
                  </div>
                </div>

                {/* Inventory table */}
                <div className="overflow-x-auto border border-[#27272a] rounded-lg">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-[#27272a] bg-[#111113]">
                        <th className="text-left px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider">
                          Material
                        </th>
                        <th className="text-center px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider w-20">
                          Type
                        </th>
                        <th className="text-right px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider w-20">
                          Pieces
                        </th>
                        <th className="text-left px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider">
                          Leftover Details
                        </th>
                        <th className="text-right px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider">
                          Unit Price
                        </th>
                        <th className="text-right px-4 py-2.5 text-xs font-medium text-zinc-500 uppercase tracking-wider">
                          Est. Value
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#27272a]/50">
                      {stockInventory.items.map((item, i) => (
                        <tr key={i} className="hover:bg-[#1c1c21] transition-colors">
                          <td className="px-4 py-2.5">
                            <div className="text-white text-xs">{item.description}</div>
                            <div className="text-zinc-500 text-[10px] font-mono">
                              {item.partNumber} ({item.finish})
                            </div>
                          </td>
                          <td className="px-4 py-2.5 text-center">
                            <span
                              className={`text-[10px] font-semibold uppercase px-1.5 py-0.5 rounded ${
                                item.type === 'profile'
                                  ? 'bg-blue-500/20 text-blue-400'
                                  : 'bg-orange-500/20 text-orange-400'
                              }`}
                            >
                              {item.type === 'profile' ? 'Profile' : 'Accessory'}
                            </span>
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono text-xs text-zinc-300 tabular-nums">
                            {item.quantity}
                          </td>
                          <td className="px-4 py-2.5">
                            <div className="text-xs text-zinc-400 font-mono">
                              {item.pieceSummary}
                            </div>
                            {item.type === 'profile' && item.totalLength > 0 && (
                              <div className="text-[10px] text-zinc-600 mt-0.5">
                                Total: {item.totalLength.toFixed(2)} ft
                              </div>
                            )}
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono text-xs text-zinc-400 tabular-nums">
                            {item.unitPrice > 0
                              ? formatCurrency(item.unitPrice)
                              : '—'}
                            {item.type === 'profile' && <span className="text-zinc-600">/ft</span>}
                            {item.type === 'accessory' && <span className="text-zinc-600">/ea</span>}
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono text-xs text-emerald-400 tabular-nums">
                            {formatCurrency(item.estimatedValue)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr className="border-t border-[#27272a] bg-[#111113]">
                        <td className="px-4 py-2.5 text-xs font-semibold text-white" colSpan={2}>
                          TOTAL
                        </td>
                        <td className="px-4 py-2.5 text-right font-mono text-xs font-semibold text-white tabular-nums">
                          {stockInventory.totalPieces}
                        </td>
                        <td className="px-4 py-2.5" colSpan={2}></td>
                        <td className="px-4 py-2.5 text-right font-mono text-xs font-semibold text-emerald-400 tabular-nums">
                          {formatCurrency(stockInventory.totalValue)}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>

                <p className="text-[10px] text-zinc-600 italic">
                  Leftover pieces are accumulated across all elevations from the bin-packing/cutting optimization.
                  These pieces may be reusable in future projects or additional elevations.
                </p>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}
