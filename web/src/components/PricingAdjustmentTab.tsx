'use client';

import { useState, useCallback, useEffect } from 'react';
import {
  Save,
  Loader2,
  Settings2,
  Info,
  DollarSign,
  TrendingUp,
  ChevronDown,
  ChevronUp,
} from 'lucide-react';
import type { ProjectSettings } from '@/types';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

interface PricingAdjustmentTabProps {
  settings: ProjectSettings;
  onSettingsUpdate: (newSettings: ProjectSettings) => Promise<void>;
}

// ---------------------------------------------------------------------------
// Defaults (matching Python's _pricing_defaults)
// ---------------------------------------------------------------------------

const PRICING_DEFAULTS = {
  discount_multiplier_low: 0.614,
  discount_multiplier_high: 0.572,
  discount_threshold: 50000,
  glass_per_sqft: 10.5,
  fabrication_cost_per_joint: 15,
};

// ---------------------------------------------------------------------------
// Shared styling
// ---------------------------------------------------------------------------

const inputClass =
  'bg-[#0c0c12] border border-[#1e1e2a] text-white rounded-xl px-3.5 py-2.5 w-full focus:outline-none focus:ring-2 focus:ring-[#3b82f6]/20 focus:border-[#3b82f6] transition-colors duration-200 text-sm';
const labelClass = 'block text-sm font-medium text-[#ffffff] mb-1.5';
const cardClass = 'bg-[#111118] border border-[#1e1e2a] rounded-2xl p-6 space-y-5 shadow-lg shadow-black/15';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------



// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export default function PricingAdjustmentTab({
  settings,
  onSettingsUpdate,
}: PricingAdjustmentTabProps) {
  // ---- Section collapse state ----
  const [collapsedSections, setCollapsedSections] = useState<Record<string, boolean>>({});
  const toggleSection = (key: string) =>
    setCollapsedSections((prev) => ({ ...prev, [key]: !prev[key] }));

  // ---- Save indicators ----
  const [savingPricing, setSavingPricing] = useState(false);
  const [savedPricing, setSavedPricing] = useState(false);
  const [savingAdditional, setSavingAdditional] = useState(false);
  const [savingMarkups, setSavingMarkups] = useState(false);

  // ---- Pricing fields state ----
  const [discountMultiplierLow, setDiscountMultiplierLow] = useState(
    settings.discount_multiplier_low ?? PRICING_DEFAULTS.discount_multiplier_low,
  );
  const [discountMultiplierHigh, setDiscountMultiplierHigh] = useState(
    settings.discount_multiplier_high ?? PRICING_DEFAULTS.discount_multiplier_high,
  );
  const [discountThreshold, setDiscountThreshold] = useState(
    settings.discount_threshold ?? PRICING_DEFAULTS.discount_threshold,
  );
  const [glassPerSqft, setGlassPerSqft] = useState(
    settings.glass_per_sqft ?? PRICING_DEFAULTS.glass_per_sqft,
  );
  const [fabricationCostPerJoint, setFabricationCostPerJoint] = useState(
    settings.fabrication_cost_per_joint ?? PRICING_DEFAULTS.fabrication_cost_per_joint,
  );

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

  // Sync from props when settings change externally
  useEffect(() => {
    setDiscountMultiplierLow(settings.discount_multiplier_low ?? PRICING_DEFAULTS.discount_multiplier_low);
    setDiscountMultiplierHigh(settings.discount_multiplier_high ?? PRICING_DEFAULTS.discount_multiplier_high);
    setDiscountThreshold(settings.discount_threshold ?? PRICING_DEFAULTS.discount_threshold);
    setGlassPerSqft(settings.glass_per_sqft ?? PRICING_DEFAULTS.glass_per_sqft);
    setFabricationCostPerJoint(settings.fabrication_cost_per_joint ?? PRICING_DEFAULTS.fabrication_cost_per_joint);
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
  }, [settings]);

  // ---- Save Pricing ----
  const handleSavePricing = useCallback(async () => {
    setSavingPricing(true);
    setSavedPricing(false);
    try {
      await onSettingsUpdate({
        ...settings,
        discount_multiplier_low: discountMultiplierLow || PRICING_DEFAULTS.discount_multiplier_low,
        discount_multiplier_high: discountMultiplierHigh || PRICING_DEFAULTS.discount_multiplier_high,
        discount_threshold: discountThreshold || PRICING_DEFAULTS.discount_threshold,
        glass_per_sqft: glassPerSqft || PRICING_DEFAULTS.glass_per_sqft,
        fabrication_cost_per_joint: fabricationCostPerJoint || PRICING_DEFAULTS.fabrication_cost_per_joint,
      });
      setSavedPricing(true);
      setTimeout(() => setSavedPricing(false), 2000);
    } finally {
      setSavingPricing(false);
    }
  }, [
    settings, onSettingsUpdate,
    discountMultiplierLow, discountMultiplierHigh, discountThreshold,
    glassPerSqft, fabricationCostPerJoint,
  ]);

  const handleResetPricing = useCallback(() => {
    setDiscountMultiplierLow(PRICING_DEFAULTS.discount_multiplier_low);
    setDiscountMultiplierHigh(PRICING_DEFAULTS.discount_multiplier_high);
    setDiscountThreshold(PRICING_DEFAULTS.discount_threshold);
    setGlassPerSqft(PRICING_DEFAULTS.glass_per_sqft);
    setFabricationCostPerJoint(PRICING_DEFAULTS.fabrication_cost_per_joint);
  }, []);

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
          <ChevronDown className="h-4 w-4 text-[#ffffff] ml-auto transition-transform duration-200" />
        ) : (
          <ChevronUp className="h-4 w-4 text-[#ffffff] ml-auto transition-transform duration-200" />
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
        <span className="absolute right-3 top-1/2 -translate-y-1/2 text-[#ffffff] text-sm">%</span>
      </div>
    </div>
  );

  return (
    <div className="mx-auto max-w-5xl space-y-6 p-6 pb-12">
      {/* ================================================================== */}
      {/* 1. Pricing Configuration (Discount Multipliers + Glass/Fab)        */}
      {/* ================================================================== */}
      <div className={cardClass}>
        <SectionHeader
          sectionKey="pricing"
          title="Pricing Configuration"
          icon={<Settings2 className="w-5 h-5 text-[#3b82f6]" />}
          action={
            <div className="flex items-center gap-2">
              <button
                onClick={handleResetPricing}
                className="px-3 py-2 text-xs font-medium text-[#ffffff] hover:text-white hover:bg-[#16161f] rounded-lg transition-colors duration-200"
              >
                Reset Defaults
              </button>
              <button
                onClick={handleSavePricing}
                disabled={savingPricing}
                className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-[#3b82f6] to-[#2563eb] hover:brightness-110 px-4 py-2 text-sm font-semibold text-white disabled:opacity-40 transition-colors duration-200"
              >
                {savingPricing ? <Loader2 className="w-4 h-4 animate-spin" /> : <Save className="w-4 h-4" />}
                {savedPricing ? 'Saved!' : 'Save'}
              </button>
            </div>
          }
        />
        {!collapsedSections.pricing && (
          <>
            {/* Info box */}
            <div className="flex items-start gap-2.5 p-4 bg-[#3b82f6]/5 border border-[#3b82f6]/10 rounded-xl">
              <Info className="w-4 h-4 text-blue-400 mt-0.5 shrink-0" />
              <p className="text-xs text-blue-300/80 leading-relaxed">
                The discount multiplier is applied only to <strong>profiles, gaskets, and accessories</strong>.
                Glass, doors, and fabrication costs are <strong>not</strong> discounted.
                The multiplier tier is determined by the total list price across all elevations.
              </p>
            </div>

            {/* Discount Multipliers */}
            <div>
              <h4 className="text-sm font-semibold text-[#c4c5d0] mb-3">Discount Multipliers</h4>
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
                <div>
                  <label className={labelClass}>
                    Multiplier (projects &lt; threshold)
                  </label>
                  <input
                    type="number"
                    className={inputClass}
                    min={0}
                    max={1}
                    step="0.001"
                    value={discountMultiplierLow}
                    onChange={(e) => setDiscountMultiplierLow(parseFloat(e.target.value) || 0)}
                  />
                  <p className="text-xs text-[#ffffff] mt-1">Default: {PRICING_DEFAULTS.discount_multiplier_low}</p>
                </div>
                <div>
                  <label className={labelClass}>
                    Multiplier (projects &ge; threshold)
                  </label>
                  <input
                    type="number"
                    className={inputClass}
                    min={0}
                    max={1}
                    step="0.001"
                    value={discountMultiplierHigh}
                    onChange={(e) => setDiscountMultiplierHigh(parseFloat(e.target.value) || 0)}
                  />
                  <p className="text-xs text-[#ffffff] mt-1">Default: {PRICING_DEFAULTS.discount_multiplier_high}</p>
                </div>
                <div>
                  <label className={labelClass}>Discount Threshold ($)</label>
                  <input
                    type="number"
                    className={inputClass}
                    min={0}
                    step="1000"
                    value={discountThreshold}
                    onChange={(e) => setDiscountThreshold(parseFloat(e.target.value) || 0)}
                  />
                  <p className="text-xs text-[#ffffff] mt-1">Default: ${PRICING_DEFAULTS.discount_threshold.toLocaleString()}</p>
                </div>
              </div>
            </div>

            {/* Glass & Fabrication */}
            <div className="border-t border-[#1e1e2a] pt-4">
              <h4 className="text-sm font-semibold text-[#c4c5d0] mb-3">Glass & Fabrication Pricing</h4>
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                <div>
                  <label className={labelClass}>Glass per sqft ($)</label>
                  <input
                    type="number"
                    className={inputClass}
                    min={0}
                    step="0.5"
                    value={glassPerSqft}
                    onChange={(e) => setGlassPerSqft(parseFloat(e.target.value) || 0)}
                  />
                  <p className="text-xs text-[#ffffff] mt-1">Default: ${PRICING_DEFAULTS.glass_per_sqft}</p>
                </div>
                <div>
                  <label className={labelClass}>Fabrication cost per joint ($)</label>
                  <input
                    type="number"
                    className={inputClass}
                    min={0}
                    step="0.5"
                    value={fabricationCostPerJoint}
                    onChange={(e) => setFabricationCostPerJoint(parseFloat(e.target.value) || 0)}
                  />
                  <p className="text-xs text-[#ffffff] mt-1">Default: ${PRICING_DEFAULTS.fabrication_cost_per_joint}</p>
                </div>
              </div>
            </div>
          </>
        )}
      </div>

      {/* ================================================================== */}
      {/* 2. Additional Cost Settings                                        */}
      {/* ================================================================== */}
      <div className={cardClass}>
        <SectionHeader
          sectionKey="additional"
          title="Additional Cost Settings"
          icon={<DollarSign className="w-5 h-5 text-[#3b82f6]" />}
          action={
            <button
              onClick={handleSaveAdditionalCosts}
              disabled={savingAdditional}
              className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-[#3b82f6] to-[#2563eb] hover:brightness-110 px-4 py-2 text-sm font-semibold text-white disabled:opacity-40 transition-colors duration-200"
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

      {/* ================================================================== */}
      {/* 3. Markup Settings                                                 */}
      {/* ================================================================== */}
      <div className={cardClass}>
        <SectionHeader
          sectionKey="markups"
          title="Markup Settings"
          icon={<TrendingUp className="w-5 h-5 text-emerald-500" />}
          action={
            <button
              onClick={handleSaveMarkups}
              disabled={savingMarkups}
              className="flex items-center gap-2 rounded-xl bg-gradient-to-r from-emerald-600 to-emerald-500 hover:brightness-110 px-4 py-2 text-sm font-semibold text-white disabled:opacity-40 transition-colors duration-200"
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


    </div>
  );
}
