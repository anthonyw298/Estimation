'use client';

import { useMemo, useState } from 'react';
import { AlertTriangle, TrendingDown, Recycle, BarChart3, Package, RotateCcw } from 'lucide-react';
import type {
  ElevationData,
  ExtraMaterial,
  ProjectSettings,
  CalculatedOutput,
  WasteMaterialBreakdown,
  WasteSuggestion,
} from '@/types';
import { getUnitPriceByPart, getPriceByPart, applyMaterialImpactInMemory, parseLengthToFeet } from '@/lib/pricing';
import { partsData } from '@/data/parts-data';

interface WasteAnalysisProps {
  elevations: Record<string, ElevationData>;
  materials: Record<string, ExtraMaterial>;
  settings: ProjectSettings;
  onResetInventory?: () => void;
}

// ---------- Shared helpers (matching export.ts / pdf-export.ts) ----------

const DISCOUNTABLE_TYPES = new Set(['profiles', 'gaskets', 'accessories']);
const GASKET_PART_NUMBERS = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

function classifyOutput(output: CalculatedOutput): string {
  const pn = output.part_number || '';
  const desc = (output.description || '').toLowerCase();
  const type = (output.type || '').toLowerCase();
  if (pn === 'GLASS_AREA' || type === 'glass') return 'glass';
  if (pn === 'JOINTS_FAB_LABOR' || type === 'joints_fab_labor' || type === 'fabrication' ||
      desc.includes('joints fabrication') || desc.includes('fabrication labor')) return 'fabrication';
  if (type === 'door' || type === 'doors') return 'doors';
  if (type === 'calculations') return 'calculations';
  if (desc.includes('gasket') || GASKET_PART_NUMBERS.has(pn)) return 'gaskets';
  if (type === 'accessory' || type === 'accessories') return 'accessories';
  return 'profiles';
}

function getMultiplier(totalListPrice: number, settings: ProjectSettings): number {
  if (settings.discount_multiplier != null) return settings.discount_multiplier;
  const threshold = settings.discount_threshold ?? 50000;
  const low = settings.discount_multiplier_low ?? 0.614;
  const high = settings.discount_multiplier_high ?? 0.572;
  return totalListPrice < threshold ? low : high;
}

function formatCurrency(value: number): string {
  return '$' + value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseKey(key: string): { partNumber: string; finish?: string } {
  const lastDashIdx = key.lastIndexOf('-');
  if (lastDashIdx > 0) {
    const possibleFinish = key.substring(lastDashIdx + 1);
    if (['clear', 'black', 'paint', 'bronze', 'mill'].includes(possibleFinish)) {
      return {
        partNumber: key.substring(0, lastDashIdx),
        finish: possibleFinish,
      };
    }
  }
  return { partNumber: key };
}

export default function WasteAnalysis({ elevations, materials, settings, onResetInventory }: WasteAnalysisProps) {
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const analysis = useMemo(() => {
    // Compute multiplier (matching export.ts / pdf-export.ts logic)
    let runningGrandTotal = 0;
    for (const elev of Object.values(elevations)) {
      if (!elev.calculated_outputs) continue;
      for (const output of elev.calculated_outputs) {
        if (output.type !== 'Calculations' && output.price != null) {
          runningGrandTotal += output.price;
        }
      }
    }
    const multiplier = getMultiplier(runningGrandTotal, settings);

    // Compute discounted total (matching export.ts / pdf-export.ts logic)
    let totalDiscountable = 0;
    let totalNonDiscountable = 0;
    for (const elev of Object.values(elevations)) {
      if (!elev.calculated_outputs) continue;
      for (const output of elev.calculated_outputs) {
        if (output.type === 'Calculations' || output.price == null) continue;
        const cat = classifyOutput(output);
        if (DISCOUNTABLE_TYPES.has(cat)) {
          totalDiscountable += output.price;
        } else {
          totalNonDiscountable += output.price;
        }
      }
    }
    const discountedTotal = (totalDiscountable * multiplier) + totalNonDiscountable;

    // ---------------------------------------------------------------
    // Aggregate all calculated_outputs across elevations by part key
    // (mirrors export.ts buildSummaryCategories Steps 2-3)
    // ---------------------------------------------------------------
    const partMap = new Map<string, {
      category: string;
      description: string;
      part_number: string;
      total_qty: number;
      quantity_list: number[];
      manual_total_cost: number;
      finish: string;
      isDiscountable: boolean;
      isManual: boolean;
    }>();

    function sumQty(qty: number | number[]): number {
      return Array.isArray(qty) ? qty.reduce((s, v) => s + Number(v), 0) : Number(qty);
    }

    for (const elev of Object.values(elevations)) {
      if (!elev.calculated_outputs) continue;
      const elevFinish = elev.finish || '';
      for (const output of elev.calculated_outputs) {
        const cat = classifyOutput(output);
        if (cat === 'calculations') continue;
        const isManual = !!(output.manual) || cat === 'glass' || cat === 'fabrication' || cat === 'doors';

        const pn = output.part_number || '';
        const isProfileOrGasket = cat === 'profiles' || cat === 'gaskets' || cat === 'glass' || cat === 'fabrication' || cat === 'doors';
        const key = isManual
          ? (pn && pn !== 'N/A'
            ? (isProfileOrGasket && elevFinish ? `MANUAL_${pn}-${elevFinish}` : `MANUAL_${pn}`)
            : `MANUAL_NO_PN_${output.description}`)
          : (cat === 'profiles' || cat === 'gaskets') && elevFinish
            ? `${pn}-${elevFinish}`
            : pn;

        const qty = sumQty(output.quantity);
        const qtyList = Array.isArray(output.quantity) ? output.quantity : [output.quantity];
        const existing = partMap.get(key);
        if (existing) {
          existing.total_qty += qty;
          existing.quantity_list = existing.quantity_list.concat(qtyList);
          if (isManual) existing.manual_total_cost += output.price ?? 0;
        } else {
          partMap.set(key, {
            category: cat,
            description: output.description,
            part_number: pn,
            total_qty: qty,
            quantity_list: [...qtyList],
            manual_total_cost: isManual ? (output.price ?? 0) : 0,
            finish: elevFinish,
            isDiscountable: DISCOUNTABLE_TYPES.has(cat),
            isManual,
          });
        }
      }
    }

    // ---------------------------------------------------------------
    // Re-price from scratch with fresh materials state so leftover
    // quantities are accurate (not consumed by cross-elevation reuse).
    // ---------------------------------------------------------------
    const freshMaterials: Record<string, ExtraMaterial> = {};

    // Build material breakdown
    const breakdown: WasteMaterialBreakdown[] = [];
    let totalWasteCost = 0;
    let totalMaterialCost = 0;

    for (const [, data] of partMap) {
      const isProfile = data.category === 'profiles';
      const isGasket = data.category === 'gaskets';
      const isAccessory = data.category === 'accessories';

      // Skip non-material categories for waste analysis
      if (!isProfile && !isGasket && !isAccessory) continue;

      let totalCost: number;
      if (data.isManual) {
        totalCost = data.manual_total_cost;
      } else {
        const useGroup = isGasket;
        let quantityForPricing: number | number[] = data.total_qty;
        if ((isProfile || isGasket) && data.part_number && data.part_number !== 'N/A') {
          const validQuantities = data.quantity_list.filter(
            (q): q is number => q != null && typeof q === 'number' && q > 0,
          );
          if (validQuantities.length > 1) {
            quantityForPricing = validQuantities;
          }
        }
        const [price, , impact] = getPriceByPart(
          data.part_number, quantityForPricing, data.finish,
          freshMaterials, false, useGroup, data.description,
        );
        if (impact) applyMaterialImpactInMemory(freshMaterials, impact);
        totalCost = price ?? 0;
      }

      const materialCost = data.isDiscountable ? totalCost * multiplier : totalCost;

      // Read leftover from the freshly computed state
      const extraKey = (isProfile || isGasket) && data.finish
        ? `${data.part_number}-${data.finish.toLowerCase()}`
        : data.part_number;
      const partState = freshMaterials[extraKey];
      const leftoverPieces = partState?.length_pieces ?? [];
      const leftoverQty = partState?.quantity ?? 0;

      const unit = (isProfile || isGasket) ? 'ft' : 'pcs';
      let wasteQty: number;
      let wasteDisplay: string;

      if (unit === 'ft') {
        wasteQty = leftoverPieces.reduce((s, l) => s + l, 0);
        wasteDisplay = wasteQty.toFixed(2) + ' ft';
      } else {
        wasteQty = leftoverQty;
        wasteDisplay = wasteQty.toFixed(0) + ' pcs';
      }

      const [unitPrice] = getUnitPriceByPart(data.part_number, data.finish);
      const wasteCost = (unitPrice ?? 0) * wasteQty * multiplier;
      const totalAcquired = data.total_qty + wasteQty;
      const wastePercentage = totalAcquired > 0 ? (wasteQty / totalAcquired) * 100 : 0;

      breakdown.push({
        part_number: data.part_number,
        description: data.description,
        finish: data.finish,
        total_quantity: data.total_qty,
        waste_quantity: wasteQty,
        waste_quantity_display: wasteDisplay,
        waste_percentage: wastePercentage,
        waste_cost: wasteCost,
        material_cost: materialCost,
        unit,
        individual_pieces: [...leftoverPieces],
      });

      totalWasteCost += wasteCost;
      totalMaterialCost += materialCost;
    }

    // Sort by waste cost descending
    breakdown.sort((a, b) => b.waste_cost - a.waste_cost);

    // Overall waste percentage: cost-based matching exports
    // Formula: residualCost / discountedTotal * 100 (same as export.ts & pdf-export.ts)
    const overallWastePercentage = discountedTotal > 0 ? (totalWasteCost / discountedTotal) * 100 : 0;

    // Generate suggestions
    const suggestions: WasteSuggestion[] = [];

    // High waste items
    const highWasteItems = breakdown.filter((m) => m.waste_percentage > 30 && m.waste_cost > 50);
    for (const item of highWasteItems.slice(0, 3)) {
      suggestions.push({
        priority: 'high',
        category: 'High Waste',
        message: `${item.part_number} (${item.finish}) has ${item.waste_percentage.toFixed(1)}% waste. Consider adjusting bay dimensions to reduce leftover material.`,
        estimated_savings: item.waste_cost * 0.5,
      });
    }

    // Medium waste items
    const mediumWasteItems = breakdown.filter(
      (m) => m.waste_percentage > 15 && m.waste_percentage <= 30 && m.waste_cost > 25
    );
    for (const item of mediumWasteItems.slice(0, 2)) {
      suggestions.push({
        priority: 'medium',
        category: 'Moderate Waste',
        message: `${item.part_number} has ${item.waste_percentage.toFixed(1)}% waste (${formatCurrency(item.waste_cost)}). Leftover pieces may be usable in future elevations.`,
        estimated_savings: item.waste_cost * 0.3,
      });
    }

    // Low: leftover reuse opportunity
    const reusableItems = breakdown.filter(
      (m) => m.individual_pieces.length > 0 && m.waste_percentage <= 15
    );
    if (reusableItems.length > 0) {
      suggestions.push({
        priority: 'low',
        category: 'Reuse Opportunity',
        message: `${reusableItems.length} material(s) have small leftover pieces that could be reused across elevations to minimize new purchases.`,
        estimated_savings: null,
      });
    }

    if (overallWastePercentage > 20) {
      suggestions.push({
        priority: 'high',
        category: 'Overall Waste',
        message: `Overall waste is ${overallWastePercentage.toFixed(1)}% (${formatCurrency(totalWasteCost)}). Review opening dimensions and bay layouts for better material utilization.`,
        estimated_savings: totalWasteCost * 0.3,
      });
    }

    return {
      breakdown,
      totalWasteCost,
      totalMaterialCost,
      overallWastePercentage,
      suggestions,
      multiplier,
      freshMaterials,
    };
  }, [elevations, settings]);

  const { breakdown, totalWasteCost, totalMaterialCost, overallWastePercentage, suggestions, multiplier, freshMaterials } = analysis;

  // ---- Stock / Leftover Inventory ----
  // Uses freshMaterials (computed from scratch, same as Material Breakdown)
  // instead of the stored materials prop, so it always matches the Excel export
  // and isn't affected by stale/corrupted inventory data.
  const stockInventory = useMemo(() => {
    const items: Array<{
      partNumber: string;
      finish: string;
      description: string;
      type: 'profile' | 'accessory';
      pieces: number[];
      totalLength: number;
      quantity: number;
      unitPrice: number;
      estimatedValue: number;
      pieceSummary: string;
    }> = [];

    let totalPieces = 0;
    let totalValue = 0;

    // Build a description lookup from the breakdown (which has correct descriptions)
    const descriptionMap = new Map<string, string>();
    for (const b of breakdown) {
      descriptionMap.set(b.part_number, b.description);
    }

    for (const [key, mat] of Object.entries(freshMaterials)) {
      const hasLengthPieces = mat.length_pieces && mat.length_pieces.length > 0;
      const hasQuantity = mat.quantity > 0;
      if (!hasLengthPieces && !hasQuantity) continue;

      const { partNumber, finish } = parseKey(key);
      const [unitPrice] = getUnitPriceByPart(partNumber, finish ?? 'clear');
      const description = descriptionMap.get(partNumber) ?? partNumber;

      if (hasLengthPieces) {
        const pieces = mat.length_pieces.filter((l) => l > 0).sort((a, b) => b - a);
        if (pieces.length === 0) continue;

        const totalLength = pieces.reduce((s, l) => s + l, 0);
        const value = (unitPrice ?? 0) * totalLength * multiplier;

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
        const value = (unitPrice ?? 0) * mat.quantity * multiplier;
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

    items.sort((a, b) => b.estimatedValue - a.estimatedValue);
    return { items, totalPieces, totalValue };
  }, [freshMaterials, breakdown, multiplier]);

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

  function getPriorityStyles(priority: 'high' | 'medium' | 'low'): string {
    switch (priority) {
      case 'high':
        return 'border-red-500/30 bg-red-500/5';
      case 'medium':
        return 'border-yellow-500/30 bg-yellow-500/5';
      case 'low':
        return 'border-emerald-500/30 bg-emerald-500/5';
    }
  }

  function getPriorityBadge(priority: 'high' | 'medium' | 'low'): string {
    switch (priority) {
      case 'high':
        return 'bg-red-500/20 text-red-400';
      case 'medium':
        return 'bg-yellow-500/20 text-yellow-400';
      case 'low':
        return 'bg-emerald-500/20 text-emerald-400';
    }
  }

  if (breakdown.length === 0) {
    return (
      <div className="bg-[#111118] border border-[#1e1e2a] rounded-2xl p-8 text-center animate-fade-up opacity-0">
        <Recycle className="w-10 h-10 text-[#2a2a3a] mx-auto mb-3" />
        <p className="text-sm text-[#ffffff]">
          No material data available. Calculate elevations to see waste analysis.
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* Overview card */}
      <div className="bg-[#111118] border border-[#1e1e2a] rounded-2xl p-6 shadow-lg shadow-black/15">
        <div className="flex items-center gap-2.5 mb-5">
          <div className="w-8 h-8 rounded-lg bg-[#3b82f6]/10 border border-[#3b82f6]/15 flex items-center justify-center">
            <BarChart3 className="w-4 h-4 text-[#3b82f6]" />
          </div>
          <h3 className="text-sm font-semibold text-[#ffffff] tracking-tight">
            Waste Overview
          </h3>
        </div>

        <div className="grid grid-cols-3 gap-3.5">
          {/* Overall waste percentage */}
          <div className="stat-card stat-card-rose bg-[#08080e] border border-[#1e1e2a] rounded-xl p-5 text-center">
            <p className="text-[10px] text-[#ffffff] mb-1.5 font-semibold uppercase tracking-wider">Overall Waste</p>
            <p className={`text-2xl font-bold font-mono tabular-nums ${getWasteColor(overallWastePercentage)}`}>
              {overallWastePercentage.toFixed(1)}%
            </p>
          </div>

          {/* Total waste cost */}
          <div className="stat-card stat-card-amber bg-[#08080e] border border-[#1e1e2a] rounded-xl p-5 text-center">
            <p className="text-[10px] text-[#ffffff] mb-1.5 font-semibold uppercase tracking-wider">Waste Cost</p>
            <p className="text-2xl font-bold font-mono text-yellow-400 tabular-nums">
              {formatCurrency(totalWasteCost)}
            </p>
          </div>

          {/* Material cost */}
          <div className="stat-card stat-card-blue bg-[#08080e] border border-[#1e1e2a] rounded-xl p-5 text-center">
            <p className="text-[10px] text-[#ffffff] mb-1.5 font-semibold uppercase tracking-wider">Material Cost</p>
            <p className="text-2xl font-bold font-mono text-[#ffffff] tabular-nums">
              {formatCurrency(totalMaterialCost)}
            </p>
          </div>
        </div>
      </div>

      {/* Material breakdown table */}
      <div className="bg-[#111118] border border-[#1e1e2a] rounded-2xl overflow-hidden shadow-lg shadow-black/15">
        <div className="px-5 py-4 border-b border-[#1e1e2a] flex items-center gap-2.5">
          <div className="w-7 h-7 rounded-lg bg-[#3b82f6]/10 border border-[#3b82f6]/15 flex items-center justify-center">
            <TrendingDown className="w-3.5 h-3.5 text-[#3b82f6]" />
          </div>
          <h3 className="text-sm font-semibold text-[#ffffff] tracking-tight">
            Material Breakdown
          </h3>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-[#1e1e2a] bg-[#0a0a10]">
                <th className="text-left px-4 py-3 text-xs font-medium text-[#ffffff] uppercase tracking-wider">
                  Part Number
                </th>
                <th className="text-left px-4 py-3 text-xs font-medium text-[#ffffff] uppercase tracking-wider">
                  Description
                </th>
                <th className="text-left px-4 py-3 text-xs font-medium text-[#ffffff] uppercase tracking-wider">
                  Finish
                </th>
                <th className="text-right px-4 py-3 text-xs font-medium text-[#ffffff] uppercase tracking-wider">
                  Used Qty
                </th>
                <th className="text-right px-4 py-3 text-xs font-medium text-[#ffffff] uppercase tracking-wider">
                  Waste Qty
                </th>
                <th className="text-right px-4 py-3 text-xs font-medium text-[#ffffff] uppercase tracking-wider w-32">
                  Waste %
                </th>
                <th className="text-right px-4 py-3 text-xs font-medium text-[#ffffff] uppercase tracking-wider">
                  Waste Cost
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#1e1e2a]">
              {breakdown.map((mat, i) => (
                <tr
                  key={`${mat.part_number}-${mat.finish}-${i}`}
                  className="table-row-hover"
                >
                  <td className="px-4 py-3 text-[#ffffff] font-mono text-xs">
                    {mat.part_number}
                  </td>
                  <td className="px-4 py-3 text-[#ffffff] text-xs max-w-[200px] truncate">
                    {mat.description}
                  </td>
                  <td className="px-4 py-3 text-[#ffffff] text-xs capitalize">
                    {mat.finish}
                  </td>
                  <td className="px-4 py-3 text-[#ffffff] text-xs text-right font-mono tabular-nums">
                    {mat.unit === 'ft' ? mat.total_quantity.toFixed(2) : mat.total_quantity.toFixed(0)}{' '}
                    <span className="text-[#ffffff]">{mat.unit}</span>
                  </td>
                  <td className="px-4 py-3 text-xs text-right font-mono tabular-nums">
                    <span className={getWasteColor(mat.waste_percentage)}>
                      {mat.waste_quantity_display}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-xs text-right">
                    <div className="flex items-center justify-end gap-2">
                      <div className="w-20 h-2 bg-[#1e1e2a] rounded-full overflow-hidden">
                        <div
                          className={`h-full rounded-full transition-all duration-500 ${getWasteBarColor(mat.waste_percentage)}`}
                          style={{
                            width: `${Math.min(mat.waste_percentage, 100)}%`,
                            boxShadow: mat.waste_percentage > 20 ? '0 0 8px rgba(248, 113, 113, 0.3)' : mat.waste_percentage > 10 ? '0 0 6px rgba(234, 179, 8, 0.2)' : 'none'
                          }}
                        />
                      </div>
                      <span className={`font-mono tabular-nums ${getWasteColor(mat.waste_percentage)}`}>
                        {mat.waste_percentage.toFixed(1)}%
                      </span>
                    </div>
                  </td>
                  <td className="px-4 py-3 text-yellow-400 text-xs text-right font-mono tabular-nums">
                    {formatCurrency(mat.waste_cost)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Optimization suggestions */}
      {suggestions.length > 0 && (
        <div className="space-y-3">
          <div className="flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 text-[#3b82f6]" />
            <h3 className="text-sm font-semibold text-[#ffffff] tracking-tight">
              Optimization Suggestions
            </h3>
          </div>

          {suggestions.map((suggestion, i) => (
            <div
              key={i}
              className={`border rounded-xl p-4 ${getPriorityStyles(suggestion.priority)}`}
            >
              <div className="flex items-start justify-between gap-3">
                <div className="flex-1">
                  <div className="flex items-center gap-2 mb-1">
                    <span
                      className={`text-[10px] font-semibold uppercase tracking-wider px-1.5 py-0.5 rounded ${getPriorityBadge(suggestion.priority)}`}
                    >
                      {suggestion.priority}
                    </span>
                    <span className="text-xs text-[#ffffff] font-medium">
                      {suggestion.category}
                    </span>
                  </div>
                  <p className="text-xs text-[#ffffff] leading-relaxed">
                    {suggestion.message}
                  </p>
                </div>
                {suggestion.estimated_savings != null && (
                  <div className="text-right shrink-0">
                    <p className="text-[10px] text-[#ffffff] uppercase font-medium">
                      Est. Savings
                    </p>
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

      {/* Stock / Leftover Inventory */}
      <div className="bg-[#111118] border border-[#1e1e2a] rounded-2xl overflow-hidden shadow-lg shadow-black/15">
        <div className="px-5 py-4 border-b border-[#1e1e2a] flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-7 h-7 rounded-lg bg-purple-500/10 border border-purple-500/15 flex items-center justify-center">
              <Package className="w-3.5 h-3.5 text-purple-400" />
            </div>
            <h3 className="text-sm font-semibold text-[#ffffff] tracking-tight">
              Stock / Leftover Inventory
            </h3>
          </div>
          {onResetInventory && stockInventory.items.length > 0 && (
            <button
              onClick={() => setShowResetConfirm(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-red-400 hover:text-red-300 hover:bg-red-500/10 rounded-xl border border-red-500/20 transition-colors duration-200"
              title="Clear all materials and mark elevations for recalculation"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Reset Inventory
            </button>
          )}
        </div>

        {stockInventory.items.length === 0 ? (
          <div className="text-center py-8 animate-fade-up">
            <Package className="w-8 h-8 text-[#2a2a3a] mx-auto mb-2" />
            <p className="text-sm text-[#ffffff]">
              No leftover inventory. Calculate elevations to generate material leftovers.
            </p>
          </div>
        ) : (
          <div className="p-5 space-y-4">
            {/* Summary stats */}
            <div className="grid grid-cols-3 gap-3.5">
              <div className="stat-card stat-card-purple bg-[#08080e] border border-[#1e1e2a] rounded-xl p-5 text-center">
                <p className="text-[10px] text-[#ffffff] mb-1.5 font-semibold uppercase tracking-wider">Total Materials</p>
                <p className="text-2xl font-bold font-mono text-purple-400 tabular-nums">
                  {stockInventory.items.length}
                </p>
                <p className="text-[10px] text-[#ffffff] mt-1 font-medium">unique part numbers</p>
              </div>
              <div className="stat-card stat-card-blue bg-[#08080e] border border-[#1e1e2a] rounded-xl p-5 text-center">
                <p className="text-[10px] text-[#ffffff] mb-1.5 font-semibold uppercase tracking-wider">Total Pieces</p>
                <p className="text-2xl font-bold font-mono text-blue-400 tabular-nums">
                  {stockInventory.totalPieces}
                </p>
                <p className="text-[10px] text-[#ffffff] mt-1 font-medium">leftover pieces/units</p>
              </div>
              <div className="stat-card stat-card-emerald bg-[#08080e] border border-[#1e1e2a] rounded-xl p-5 text-center">
                <p className="text-[10px] text-[#ffffff] mb-1.5 font-semibold uppercase tracking-wider">Estimated Value</p>
                <p className="text-2xl font-bold font-mono text-emerald-400 tabular-nums">
                  {formatCurrency(stockInventory.totalValue)}
                </p>
                <p className="text-[10px] text-[#ffffff] mt-1 font-medium">reusable material value</p>
              </div>
            </div>

            {/* Inventory table */}
            <div className="overflow-x-auto border border-[#1e1e2a] rounded-lg overflow-hidden">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-[#1e1e2a] bg-[#0a0a10]">
                    <th className="text-left px-4 py-2.5 text-xs font-medium text-[#ffffff] uppercase tracking-wider">Material</th>
                    <th className="text-center px-4 py-2.5 text-xs font-medium text-[#ffffff] uppercase tracking-wider w-20">Type</th>
                    <th className="text-right px-4 py-2.5 text-xs font-medium text-[#ffffff] uppercase tracking-wider w-20">Pieces</th>
                    <th className="text-left px-4 py-2.5 text-xs font-medium text-[#ffffff] uppercase tracking-wider">Leftover Details</th>
                    <th className="text-right px-4 py-2.5 text-xs font-medium text-[#ffffff] uppercase tracking-wider">Unit Price</th>
                    <th className="text-right px-4 py-2.5 text-xs font-medium text-[#ffffff] uppercase tracking-wider">Est. Value</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#1e1e2a]">
                  {stockInventory.items.map((item, i) => (
                    <tr key={i} className="table-row-hover">
                      <td className="px-4 py-2.5">
                        <div className="text-[#ffffff] text-xs">{item.description}</div>
                        <div className="text-[#ffffff] text-[10px] font-mono">{item.partNumber}{item.type === 'profile' ? ` (${item.finish})` : ''}</div>
                      </td>
                      <td className="px-4 py-2.5 text-center">
                        <span className={`text-[10px] font-semibold uppercase px-1.5 py-0.5 rounded ${
                          item.type === 'profile'
                            ? 'bg-blue-500/15 text-blue-400'
                            : 'bg-orange-500/15 text-orange-400'
                        }`}>
                          {item.type === 'profile' ? 'Profile' : 'Accessory'}
                        </span>
                      </td>
                      <td className="px-4 py-2.5 text-right font-mono text-xs text-[#ffffff] tabular-nums">
                        {item.quantity}
                      </td>
                      <td className="px-4 py-2.5">
                        <div className="text-xs text-[#ffffff] font-mono">{item.pieceSummary}</div>
                        {item.type === 'profile' && item.totalLength > 0 && (
                          <div className="text-[10px] text-[#ffffff] mt-0.5">
                            Total: {item.totalLength.toFixed(2)} ft
                          </div>
                        )}
                      </td>
                      <td className="px-4 py-2.5 text-right font-mono text-xs text-[#ffffff] tabular-nums">
                        {item.unitPrice > 0 ? formatCurrency(item.unitPrice) : '\u2014'}
                        {item.type === 'profile' && <span className="text-[#ffffff]">/ft</span>}
                        {item.type === 'accessory' && <span className="text-[#ffffff]">/ea</span>}
                      </td>
                      <td className="px-4 py-2.5 text-right font-mono text-xs text-emerald-400 tabular-nums">
                        {formatCurrency(item.estimatedValue)}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="border-t border-[#1e1e2a] bg-[#0a0a10]">
                    <td className="px-4 py-2.5 text-xs font-semibold text-[#ffffff]" colSpan={2}>TOTAL</td>
                    <td className="px-4 py-2.5 text-right font-mono text-xs font-semibold text-[#ffffff] tabular-nums">
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

            <p className="text-[10px] text-[#ffffff] italic">
              Leftover pieces are accumulated across all elevations from the bin-packing/cutting optimization.
              These pieces may be reusable in future projects or additional elevations.
            </p>
          </div>
        )}
      </div>
      {/* Reset Inventory Confirmation Modal */}
      {showResetConfirm && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-md animate-overlay"
          onClick={() => setShowResetConfirm(false)}
        >
          <div
            className="bg-[#111118] border border-[#1e1e2a] rounded-2xl w-full max-w-sm mx-4 p-7 shadow-2xl shadow-black/60 animate-scale-in"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-full bg-red-500/10 flex items-center justify-center flex-shrink-0">
                <RotateCcw className="w-5 h-5 text-red-400" />
              </div>
              <div>
                <h3 className="text-base font-semibold text-[#ffffff]">
                  Reset Inventory
                </h3>
                <p className="text-xs text-[#ffffff]">
                  This will require recalculating all elevations
                </p>
              </div>
            </div>

            <p className="text-sm text-[#ffffff] mb-4">
              This will clear all leftover inventory and material tracking data.
              Elevation prices and exports will continue to work normally.
            </p>
            <p className="text-xs text-yellow-400/80 mb-6">
              Recalculate each elevation when convenient to rebuild accurate material tracking and waste data.
            </p>

            <div className="flex items-center gap-3 justify-end">
              <button
                onClick={() => setShowResetConfirm(false)}
                className="px-4 py-2 text-sm font-medium text-[#ffffff] hover:text-[#ffffff] rounded-lg hover:bg-[#1e1e2a] transition-colors duration-200"
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  setShowResetConfirm(false);
                  onResetInventory?.();
                }}
                className="flex items-center gap-2 px-4 py-2 bg-red-500 hover:bg-red-600 text-white text-sm font-medium rounded-lg transition-colors duration-200"
              >
                <RotateCcw className="w-4 h-4" />
                Reset & Clear
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
