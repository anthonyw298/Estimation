'use client';

import { useMemo } from 'react';
import type { ElevationData, ExtraMaterial, ProjectSettings } from '@/types';
import {
  getPriceByPart,
  getUnitPriceByPart,
  applyMaterialImpactInMemory,
} from '@/lib/pricing';
import { PART_NUMBER_MAP } from '@/data/part-number';

interface CostSummaryProps {
  elevations: Record<string, ElevationData>;
  materials: Record<string, ExtraMaterial>;
  settings: ProjectSettings;
}

function formatCurrency(value: number): string {
  return '$' + value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Types that get discount multiplier applied
const DISCOUNTABLE_TYPES = new Set(['profiles', 'accessories', 'gaskets']);
// Types that do NOT get discounted
// 'Glass', 'Fabrication', 'Doors', 'Calculations' are NOT discounted

const GASKET_PARTS = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

export default function CostSummary({ elevations, materials, settings }: CostSummaryProps) {
  const summary = useMemo(() => {
    // ----------------------------------------------------------------
    // Step 1: Aggregate quantities across all elevations by part+finish
    // This mirrors the Python excel_generator create_summary_sheet logic.
    // ----------------------------------------------------------------
    const profileKeys = PART_NUMBER_MAP['profiles']
      ? Object.keys(PART_NUMBER_MAP['profiles'])
      : [];
    const accessoryKeys = PART_NUMBER_MAP['accessories']
      ? Object.keys(PART_NUMBER_MAP['accessories'])
      : [];

    interface AggregatedItem {
      key: string;
      partNumber: string;
      description: string;
      finish: string;
      quantityTotal: number;
      quantityList: number[];
      type: string; // 'profiles' | 'accessories' | 'gaskets' | 'Glass' | 'Fabrication' | 'Doors'
      manual: boolean;
      manualPrice: number; // per-unit price for manual items (glass, fab, doors)
      unit: string;
    }

    const aggregatedMap = new Map<string, AggregatedItem>();

    for (const [, elev] of Object.entries(elevations)) {
      if (!elev.calculated_outputs || elev.calculated_outputs.length === 0) continue;
      const elevFinish = (elev.finish || 'clear').toLowerCase();

      for (const output of elev.calculated_outputs) {
        const pn = output.part_number?.trim() || '';
        const desc = output.description?.trim() || '';
        const manual = output.manual ?? false;
        const qty = output.quantity;
        const qtyNum = Array.isArray(qty)
          ? qty.reduce((s: number, v: number) => s + Number(v), 0)
          : Number(qty) || 0;
        const qtyList: number[] = Array.isArray(qty)
          ? qty.map(Number)
          : [Number(qty) || 0];

        if (output.type === 'Calculations') continue;

        const isProfile = profileKeys.includes(pn);
        const isGasket = desc.toLowerCase().includes('gasket') || GASKET_PARTS.has(pn);
        const isAccessory = accessoryKeys.includes(pn) || output.type === 'accessories';
        const isGlass = pn === 'GLASS_AREA' || output.type === 'Glass';
        const isFab = output.type === 'Fabrication' || desc.toLowerCase().includes('fabrication');
        const isDoor = output.type === 'Doors' || output.type === 'Door';

        let itemType = output.type || 'accessories';
        if (isProfile) itemType = 'profiles';
        else if (isGasket) itemType = 'gaskets';
        else if (isAccessory) itemType = 'accessories';
        else if (isGlass) itemType = 'Glass';
        else if (isFab) itemType = 'Fabrication';
        else if (isDoor) itemType = 'Doors';

        // Build aggregation key
        const needsFinish = isProfile || isGasket;
        const key = (manual || isGlass || isFab || isDoor)
          ? `MANUAL_${pn || desc}_${needsFinish ? elevFinish : ''}`
          : `${pn}${needsFinish ? `-${elevFinish}` : ''}`;

        const existing = aggregatedMap.get(key);
        if (existing) {
          existing.quantityTotal += qtyNum;
          existing.quantityList.push(...qtyList);
          // For manual items, re-average the price
          if (existing.manual) {
            const oldCost = existing.manualPrice * (existing.quantityTotal - qtyNum);
            const newCost = (output.price ?? 0);
            // For manual items stored with total price, we need per-unit
            // Doors store total, glass/fab store total, etc.
            const perUnit = qtyNum > 0 ? newCost / qtyNum : 0;
            existing.manualPrice = existing.quantityTotal > 0
              ? (oldCost + (perUnit * qtyNum)) / existing.quantityTotal
              : 0;
          }
        } else {
          const perUnit = (manual || isGlass || isFab || isDoor)
            ? (qtyNum > 0 ? (output.price ?? 0) / qtyNum : output.price ?? 0)
            : 0;
          aggregatedMap.set(key, {
            key,
            partNumber: pn,
            description: desc,
            finish: elevFinish,
            quantityTotal: qtyNum,
            quantityList: [...qtyList],
            type: itemType,
            manual: manual || isGlass || isFab || isDoor,
            manualPrice: perUnit,
            unit: output.unit || (isProfile || isGasket ? 'ft' : isAccessory ? 'pcs' : 'pcs'),
          });
        }
      }
    }

    // ----------------------------------------------------------------
    // Step 2: Price aggregated items from scratch with fresh state
    // Matches Python: summary=False with empty summary_extra_materials_state
    // ----------------------------------------------------------------
    const freshMaterials: Record<string, ExtraMaterial> = {};
    let totalDiscountable = 0;
    let totalNonDiscountable = 0;

    for (const [, item] of aggregatedMap) {
      if (item.type === 'Calculations') continue;

      let itemCost = 0;

      if (item.manual) {
        // Manual items: glass, fabrication, doors
        itemCost = item.manualPrice * item.quantityTotal;
      } else {
        // Standard parts: reprice from scratch with fresh materials
        const isProfile = profileKeys.includes(item.partNumber);
        const isGasket = item.description.toLowerCase().includes('gasket') || GASKET_PARTS.has(item.partNumber);
        const useGroup = isProfile || isGasket;

        // For profiles/gaskets, pass quantity list for optimal cutting
        const qtyForPricing = (isProfile || isGasket) && item.quantityList.length > 1
          ? item.quantityList.filter(q => q > 0)
          : item.quantityTotal;

        const [price, , impact] = getPriceByPart(
          item.partNumber,
          qtyForPricing,
          item.finish,
          freshMaterials,
          false, // summary=false: track materials to accumulate leftovers across items
          useGroup,
          item.description,
        );

        if (impact) {
          applyMaterialImpactInMemory(freshMaterials, impact);
        }

        itemCost = price ?? 0;
      }

      // Categorize into discountable vs non-discountable
      if (DISCOUNTABLE_TYPES.has(item.type)) {
        totalDiscountable += itemCost;
      } else {
        totalNonDiscountable += itemCost;
      }
    }

    const totalListPrice = totalDiscountable + totalNonDiscountable;

    // ----------------------------------------------------------------
    // Per-elevation display costs (for the breakdown list)
    // These use the stored per-elevation prices for informational display.
    // ----------------------------------------------------------------
    const elevationCosts: {
      name: string;
      listCost: number;
    }[] = [];

    for (const [name, elev] of Object.entries(elevations)) {
      if (!elev.calculated_outputs || elev.calculated_outputs.length === 0) continue;
      let cost = 0;
      for (const output of elev.calculated_outputs) {
        if (output.price == null || output.type === 'Calculations') continue;
        cost += output.price;
      }
      elevationCosts.push({ name, listCost: cost });
    }

    // Determine discount multiplier tier based on total list price
    const threshold = settings.discount_threshold ?? 50000;
    const lowMultiplier = settings.discount_multiplier_low ?? 0.614;
    const highMultiplier = settings.discount_multiplier_high ?? 0.572;

    // Use explicit discount_multiplier if set, otherwise determine by tier
    const multiplier = settings.discount_multiplier != null
      ? settings.discount_multiplier
      : totalListPrice < threshold ? lowMultiplier : highMultiplier;

    // Apply discount ONLY to discountable types (profiles, gaskets, accessories)
    const discountedTotal = (totalDiscountable * multiplier) + totalNonDiscountable;

    // Waste cost: sum of leftover pieces * unit prices * multiplier
    // (waste only comes from profiles/gaskets/accessories, so it gets multiplier)
    let wasteCost = 0;
    for (const [key, mat] of Object.entries(materials)) {
      if (!mat.length_pieces || mat.length_pieces.length === 0) continue;

      let partNumber = key;
      let finish: string | undefined;
      const lastDashIdx = key.lastIndexOf('-');
      if (lastDashIdx > 0) {
        const possibleFinish = key.substring(lastDashIdx + 1);
        if (['clear', 'black', 'paint', 'bronze', 'mill'].includes(possibleFinish)) {
          partNumber = key.substring(0, lastDashIdx);
          finish = possibleFinish;
        }
      }

      const [unitPrice] = getUnitPriceByPart(partNumber, finish);
      if (unitPrice != null) {
        const totalLeftoverLength = mat.length_pieces.reduce((sum, l) => sum + l, 0);
        wasteCost += totalLeftoverLength * unitPrice;
      }
    }

    const estimatedWasteCost = wasteCost * multiplier;

    // Additional costs
    const additionalCostPcts = [
      settings.overhead_materials_pct ?? 0,
      settings.overhead_labor_pct ?? 0,
      settings.admin_management_pct ?? 0,
      settings.engineering_pct ?? 0,
      settings.packaging_materials_pct ?? 0,
      settings.shipping_transport_pct ?? 0,
      settings.commissions_pct ?? 0,
    ];
    const totalAdditionalCostPct = additionalCostPcts.reduce((s, v) => s + v, 0);
    const additionalCostsAmount = discountedTotal * (totalAdditionalCostPct / 100);

    // Markups
    const markupPcts = [
      settings.profit_on_material_pct ?? 0,
      settings.profit_on_waste_pct ?? 0,
      settings.profit_on_glass_pct ?? 0,
      settings.profit_on_wages_pct ?? 0,
      settings.planning_technical_pct ?? 0,
      settings.commission_pct ?? 0,
    ];
    const totalMarkupPct = markupPcts.reduce((s, v) => s + v, 0);
    const markupsAmount = discountedTotal * (totalMarkupPct / 100);

    const grandTotal = discountedTotal + estimatedWasteCost + additionalCostsAmount + markupsAmount;

    return {
      elevationCosts,
      totalListPrice,
      totalDiscountable,
      totalNonDiscountable,
      multiplier,
      threshold,
      discountedTotal,
      estimatedWasteCost,
      additionalCostsAmount,
      totalAdditionalCostPct,
      markupsAmount,
      totalMarkupPct,
      grandTotal,
    };
  }, [elevations, materials, settings]);

  const {
    elevationCosts, totalListPrice, totalDiscountable, totalNonDiscountable,
    multiplier, threshold, discountedTotal, estimatedWasteCost,
    additionalCostsAmount, totalAdditionalCostPct,
    markupsAmount, totalMarkupPct, grandTotal,
  } = summary;

  if (elevationCosts.length === 0) {
    return (
      <div className="bg-[#111118] border border-[#1e1e2a] rounded-xl p-6 text-center">
        <p className="text-sm text-[#55566a]">
          No calculated elevations yet. Calculate an elevation to see the cost summary.
        </p>
      </div>
    );
  }

  return (
    <div className="bg-[#111118] border border-[#1e1e2a] rounded-xl overflow-hidden">
      {/* Header */}
      <div className="px-5 py-4 border-b border-[#1e1e2a]">
        <h3 className="text-sm font-semibold text-[#eeeef2] tracking-tight">
          Cost Summary
        </h3>
      </div>

      <div className="p-5 space-y-3">
        {/* Per-elevation breakdown */}
        {elevationCosts.map((elev) => (
          <div key={elev.name} className="flex items-center justify-between text-sm rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
            <span className="text-[#8b8d9a] truncate mr-4">{elev.name}</span>
            <span className="text-[#eeeef2] font-mono tabular-nums">
              {formatCurrency(elev.listCost)}
            </span>
          </div>
        ))}

        {/* Divider */}
        <div className="border-t border-[#1e1e2a] pt-3 space-y-2">
          <div className="flex items-center justify-between text-sm rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
            <span className="text-[#8b8d9a] font-medium">List Price Total</span>
            <span className="text-[#eeeef2] font-mono font-medium tabular-nums">
              {formatCurrency(totalListPrice)}
            </span>
          </div>
          <div className="flex items-center justify-between text-xs rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
            <span className="text-[#3e3f4d] ml-2">Discountable (profiles/gaskets/accessories)</span>
            <span className="text-[#55566a] font-mono tabular-nums">
              {formatCurrency(totalDiscountable)}
            </span>
          </div>
          <div className="flex items-center justify-between text-xs rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
            <span className="text-[#3e3f4d] ml-2">Non-discountable (glass/doors/fabrication)</span>
            <span className="text-[#55566a] font-mono tabular-nums">
              {formatCurrency(totalNonDiscountable)}
            </span>
          </div>
        </div>

        {/* Discount multiplier */}
        <div className="flex items-center justify-between text-sm rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
          <span className="text-[#8b8d9a]">
            Discount Multiplier
            <span className="text-xs text-[#3e3f4d] ml-1">
              ({totalListPrice < threshold ? `<$${(threshold/1000).toFixed(0)}k` : `>=$${(threshold/1000).toFixed(0)}k`})
            </span>
          </span>
          <span className="text-[#3b82f6] font-mono font-semibold tabular-nums">
            x {multiplier.toFixed(3)}
          </span>
        </div>

        {/* Discounted Total */}
        <div className="flex items-center justify-between text-sm rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
          <span className="text-[#8b8d9a] font-medium">Discounted Total</span>
          <span className="text-[#eeeef2] font-mono font-medium tabular-nums">
            {formatCurrency(discountedTotal)}
          </span>
        </div>

        {/* Waste cost */}
        <div className="flex items-center justify-between text-sm rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
          <span className="text-[#8b8d9a]">Residual / Waste Cost</span>
          <span className="text-yellow-400 font-mono tabular-nums">
            {formatCurrency(estimatedWasteCost)}
          </span>
        </div>

        {/* Additional Costs */}
        {totalAdditionalCostPct > 0 && (
          <div className="flex items-center justify-between text-sm rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
            <span className="text-[#8b8d9a]">
              Additional Costs
              <span className="text-xs text-[#3e3f4d] ml-1">({totalAdditionalCostPct.toFixed(1)}%)</span>
            </span>
            <span className="text-orange-400 font-mono tabular-nums">
              {formatCurrency(additionalCostsAmount)}
            </span>
          </div>
        )}

        {/* Markups */}
        {totalMarkupPct > 0 && (
          <div className="flex items-center justify-between text-sm rounded-md px-2 py-1 hover:bg-[#0c0c12] transition-colors">
            <span className="text-[#8b8d9a]">
              Markups
              <span className="text-xs text-[#3e3f4d] ml-1">({totalMarkupPct.toFixed(1)}%)</span>
            </span>
            <span className="text-purple-400 font-mono tabular-nums">
              {formatCurrency(markupsAmount)}
            </span>
          </div>
        )}

        {/* Grand Total */}
        <div className="border-t border-[#1e1e2a] pt-3">
          <div className="flex items-center justify-between rounded-lg px-3 py-2.5 bg-gradient-to-r from-[#3b82f6]/10 via-[#111118] to-[#3b82f6]/5">
            <span className="text-[#eeeef2] font-semibold text-sm">Project Total</span>
            <span className="text-[#3b82f6] font-mono font-bold text-lg tabular-nums">
              {formatCurrency(grandTotal)}
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}
