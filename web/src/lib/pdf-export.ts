import type {
  ElevationData,
  ProjectSettings,
  ExtraMaterial,
  CalculatedOutput,
  DoorConfig,
  ReportConfig,
} from '@/types';
import { getUnitPriceByPart, getPriceByPart, applyMaterialImpactInMemory } from '@/lib/pricing';
import { upgradeGlassOutputs } from '@/lib/export';

// ---------------------------------------------------------------------------
// Helpers (shared with Excel export logic)
// ---------------------------------------------------------------------------

const DISCOUNTABLE_TYPES = new Set(['profiles', 'gaskets', 'accessories']);
const GASKET_PART_NUMBERS = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

function sumQty(qty: number | number[]): number {
  return Array.isArray(qty) ? qty.reduce((s, v) => s + Number(v), 0) : Number(qty);
}

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

function fmtCurrency(value: number): string {
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

// ---------------------------------------------------------------------------
// Pie Chart Generation (Canvas-based for PDF embedding)
// ---------------------------------------------------------------------------

function createPdfPieChart(
  materialCost: number,
  miscCost: number,
  markupCost: number,
  residualCost: number,
  fieldCost: number = 0,
): string | null {
  const grandTotal = materialCost + miscCost + markupCost + residualCost + fieldCost;
  if (grandTotal <= 0) return null;

  const chartWidth = 420;
  const chartHeight = 400;
  const centerX = chartWidth / 2;
  const centerY = 160;
  const radius = 90;

  const canvas = document.createElement('canvas');
  canvas.width = chartWidth;
  canvas.height = chartHeight;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;

  // White background
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, chartWidth, chartHeight);

  const materialPct = (materialCost / grandTotal) * 100;
  const miscPct = (miscCost / grandTotal) * 100;
  const markupPct = (markupCost / grandTotal) * 100;
  const residualPct = (residualCost / grandTotal) * 100;
  const fieldPct = (fieldCost / grandTotal) * 100;

  const MATERIAL_COLOR = '#4472C4';
  const MISC_COLOR = '#548235';
  const MARKUP_COLOR = '#7030A0';
  const RESIDUAL_COLOR = '#ED7D31';
  const FIELD_COLOR = '#BF8F00';

  interface Seg { name: string; value: number; pct: number; color: string }
  const segments: Seg[] = [];
  if (materialCost > 0) segments.push({ name: 'Active Materials', value: materialCost, pct: materialPct, color: MATERIAL_COLOR });
  if (miscCost > 0) segments.push({ name: 'Additional', value: miscCost, pct: miscPct, color: MISC_COLOR });
  if (markupCost > 0) segments.push({ name: 'Profit/Markups', value: markupCost, pct: markupPct, color: MARKUP_COLOR });
  if (residualCost > 0) segments.push({ name: 'Residual/Waste', value: residualCost, pct: residualPct, color: RESIDUAL_COLOR });
  if (fieldCost > 0) segments.push({ name: 'Field Costs', value: fieldCost, pct: fieldPct, color: FIELD_COLOR });

  // Title
  ctx.fillStyle = '#333333';
  ctx.font = 'bold 14px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('Project Cost Breakdown', centerX, 15);

  // Pie slices
  let startAngle = -Math.PI / 2;
  for (const seg of segments) {
    if (seg.pct <= 0) continue;
    const sweepAngle = (seg.pct / 100) * 2 * Math.PI;
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, startAngle, startAngle + sweepAngle);
    ctx.closePath();
    ctx.fillStyle = seg.color;
    ctx.fill();
    ctx.strokeStyle = '#FFFFFF';
    ctx.lineWidth = 2;
    ctx.stroke();
    startAngle += sweepAngle;
  }

  // Legend
  const legendY = 270;
  const legendBoxSize = 12;
  const legendSpacing = 22;
  const legendItems: Seg[] = [
    { name: 'Active Materials', value: materialCost, pct: materialPct, color: MATERIAL_COLOR },
    { name: 'Additional', value: miscCost, pct: miscPct, color: MISC_COLOR },
    { name: 'Profit/Markups', value: markupCost, pct: markupPct, color: MARKUP_COLOR },
    { name: 'Residual/Waste', value: residualCost, pct: residualPct, color: RESIDUAL_COLOR },
  ];
  if (fieldCost > 0) {
    legendItems.push({ name: 'Field Costs', value: fieldCost, pct: fieldPct, color: FIELD_COLOR });
  }

  for (let i = 0; i < legendItems.length; i++) {
    const item = legendItems[i];
    const yPos = legendY + (i * legendSpacing);
    ctx.fillStyle = item.color;
    ctx.fillRect(30, yPos, legendBoxSize, legendBoxSize);
    ctx.strokeStyle = '#333333';
    ctx.lineWidth = 1;
    ctx.strokeRect(30, yPos, legendBoxSize, legendBoxSize);
    ctx.fillStyle = '#333333';
    ctx.font = '10px Arial, sans-serif';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillText(
      `${item.name}: $${item.value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${item.pct.toFixed(1)}%)`,
      50 + legendBoxSize,
      yPos + legendBoxSize / 2,
    );
  }

  // Grand total
  ctx.fillStyle = '#333333';
  ctx.font = '9px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(
    `Grand Total: $${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    centerX,
    chartHeight - 15,
  );

  return canvas.toDataURL('image/png');
}

// ---------------------------------------------------------------------------
// PDF Export
// ---------------------------------------------------------------------------

export async function exportToPdf(
  projectName: string,
  elevations: Record<string, ElevationData>,
  doors: Record<string, DoorConfig[]>,
  settings: ProjectSettings,
  materials: Record<string, ExtraMaterial>,
  reportConfig?: ReportConfig,
): Promise<void> {
  // Dynamic import for browser
  const { default: jsPDF } = await import('jspdf');
  const autoTable = (await import('jspdf-autotable')).default;

  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'letter' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 15;

  // Pre-compute multiplier — re-price standard parts from scratch for accuracy
  let runningGrandTotal = 0;
  for (const elev of Object.values(elevations)) {
    if (!elev.calculated_outputs) continue;
    const elevFinish = elev.finish || '';
    const prePassOutputs = upgradeGlassOutputs(elev.calculated_outputs, elev, settings.glass_per_sqft ?? 10.5);
    for (const output of prePassOutputs) {
      const cat = classifyOutput(output);
      if (cat === 'calculations') continue;
      if (output.manual || cat === 'glass' || cat === 'fabrication' || cat === 'doors') {
        // Use live settings rates for glass/fab
        const qty = sumQty(output.quantity);
        if (cat === 'glass') {
          const glassRate = settings.glass_per_sqft ?? 10.5;
          runningGrandTotal += output.area_sqft != null ? qty * output.area_sqft * glassRate : qty * glassRate;
        } else if (cat === 'fabrication') {
          runningGrandTotal += qty * (settings.fabrication_cost_per_joint ?? 15.0);
        } else {
          runningGrandTotal += output.price ?? 0;
        }
      } else {
        const [price] = getPriceByPart(
          output.part_number, output.quantity, elevFinish,
          null, true, false, output.description,
        );
        runningGrandTotal += price ?? 0;
      }
    }
  }
  const multiplier = getMultiplier(runningGrandTotal, settings);

  // ---- Title Page ----
  doc.setFontSize(24);
  doc.setFont('helvetica', 'bold');
  doc.text('Cost Estimation Report', pageW / 2, 40, { align: 'center' });

  doc.setFontSize(16);
  doc.setFont('helvetica', 'normal');
  doc.text(projectName, pageW / 2, 55, { align: 'center' });

  doc.setFontSize(10);
  doc.setTextColor(128);
  doc.text(`Generated: ${new Date().toLocaleDateString()}`, pageW / 2, 68, { align: 'center' });
  doc.text(`Discount Multiplier: x${multiplier.toFixed(3)}`, pageW / 2, 75, { align: 'center' });
  doc.text('United Glass Ventures', pageW / 2, 85, { align: 'center' });
  doc.setTextColor(0);

  // ---- Helpers for autoTable Y tracking ----
  function getLastY(): number {
    return (doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? 30;
  }

  // Singular category label mapping (matches Excel's title_mapping)
  const SINGULAR_MAP: Record<string, string> = {
    profiles: 'Profile', accessories: 'Accessory', gaskets: 'Gasket',
    doors: 'Door', glass: 'Glass', fabrication: 'Labor',
  };

  // ---- Per-Elevation Pages ----
  const sortedNames = Object.keys(elevations).sort();

  for (const elevName of sortedNames) {
    const elev = elevations[elevName];
    if (!elev.calculated_outputs || elev.calculated_outputs.length === 0) continue;

    doc.addPage();
    const elevDoors = doors[elevName] || [];
    const totalCount = elev.total_count || 1;
    const showPerElev = totalCount > 1;

    // Elevation header
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text(elevName, margin, 20);

    // System info
    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    const info = [
      `System: ${elev.door_only ? 'Door Only' : elev.system_type}`,
      `Finish: ${elev.finish}`,
      `Dimensions: ${elev.opening_width_inches}" x ${elev.opening_height_inches}"`,
      `Bays: ${elev.bays_wide}W x ${elev.bays_tall}T`,
      `Count: ${totalCount}`,
      `Doors: ${elevDoors.length > 0 ? elevDoors.map(d => `${d.count}x ${d.size}`).join(', ') : 'None'}`,
    ];
    doc.text(info.join('  |  '), margin, 28);

    // Build per-category tables matching canonical headers
    const catOrder: [string, string][] = [
      ['profiles', 'PROFILES'], ['accessories', 'ACCESSORIES'], ['gaskets', 'GASKETS'],
      ['doors', 'DOORS'], ['glass', 'GLASS'], ['fabrication', 'LABOR'],
    ];

    // Per-elevation section config from report options
    const elevSections = reportConfig?.per_elevation_sections?.[elevName];

    // Build single-elevation price lookup (count=1, no residual)
    const singleElevMap = new Map<string, { price: number; quantity: number | number[] }>();
    if (showPerElev && elev.single_elevation_outputs) {
      for (const sOut of elev.single_elevation_outputs) {
        const sCat = classifyOutput(sOut);
        if (sCat === 'calculations') continue;
        const sKey = `${sCat}|${sOut.description}|${sOut.part_number}`;
        singleElevMap.set(sKey, { price: sOut.price ?? 0, quantity: sOut.quantity });
      }
    }

    // Track per-category discounted totals for elevation cost summary
    const elevCatTotals: Record<string, number> = {};
    // Track per-category single-elev discounted totals for cost summary per-elev column
    const elevCatPerElevTotals: Record<string, number> = {};
    let currentY = 33;

    // Per-elevation fresh state for inventory tracking (matches Excel buildElevationCategories)
    const elevMaterialsState: Record<string, ExtraMaterial> = {};
    const elevFinish = elev.finish || '';

    const upgradedElevOutputs = upgradeGlassOutputs(elev.calculated_outputs, elev, settings.glass_per_sqft ?? 10.5);

    for (const [catKey, catTitle] of catOrder) {
      // Skip section if unchecked in stock list
      if (elevSections?.[catKey] === false) continue;
      const items = upgradedElevOutputs.filter(o => classifyOutput(o) === catKey);
      if (items.length === 0) continue;

      const isDisc = DISCOUNTABLE_TYPES.has(catKey);

      // Build headers
      const headers: string[] = ['Description', 'Part Number', 'Total Quantity Required'];
      if (showPerElev) headers.push('Quantity Per Elevation');
      headers.push('Total List Cost');
      if (showPerElev) headers.push('Total List Cost Per Elevation');
      headers.push('Discounted Total List Cost');
      if (showPerElev) headers.push('Discounted Total List Cost Per Elevation');

      // Build data rows
      const rows: string[][] = [];
      let catOrigTotal = 0;
      let catDiscTotal = 0;
      let catOrigPerElev = 0;
      let catDiscPerElev = 0;

      for (const item of items) {
        const qty = sumQty(item.quantity);

        // Re-price standard parts from scratch (matches Excel buildElevationCategories)
        let cost: number;
        if (item.manual || catKey === 'glass' || catKey === 'fabrication' || catKey === 'doors') {
          // Use live settings rates for glass/fab
          if (catKey === 'glass') {
            const glassRate = settings.glass_per_sqft ?? 10.5;
            cost = item.area_sqft != null ? qty * item.area_sqft * glassRate : qty * glassRate;
          } else if (catKey === 'fabrication') {
            cost = qty * (settings.fabrication_cost_per_joint ?? 15.0);
          } else {
            cost = item.price ?? 0;
          }
        } else {
          const isGasket = catKey === 'gaskets';
          const isProfile = catKey === 'profiles';
          const useGroup = isProfile || isGasket;
          const shouldGroup = useGroup && Array.isArray(item.quantity) && item.quantity.length > 1;

          if (shouldGroup) {
            const [price, , impact] = getPriceByPart(
              item.part_number, item.quantity, elevFinish,
              elevMaterialsState, false, useGroup, item.description,
            );
            if (impact) applyMaterialImpactInMemory(elevMaterialsState, impact);
            cost = price ?? 0;
          } else {
            let itemTotal = 0;
            const quantities = Array.isArray(item.quantity) ? item.quantity : [item.quantity];
            for (const singleQty of quantities) {
              const [price, , impact] = getPriceByPart(
                item.part_number, singleQty, elevFinish,
                elevMaterialsState, false, useGroup, item.description,
              );
              if (impact) applyMaterialImpactInMemory(elevMaterialsState, impact);
              itemTotal += price ?? 0;
            }
            cost = itemTotal;
          }
        }

        const discounted = isDisc ? cost * multiplier : cost;
        catOrigTotal += cost;
        catDiscTotal += discounted;

        // Per-elevation: use single-elev data when available
        const sKey = `${catKey}|${item.description}|${item.part_number}`;
        const sData = singleElevMap.get(sKey);
        let perElevQty: number;
        let perElevCost: number;
        let perElevDisc: number;
        if (sData) {
          perElevQty = sumQty(sData.quantity);
          perElevCost = sData.price;
          perElevDisc = isDisc ? sData.price * multiplier : sData.price;
        } else {
          perElevQty = qty / totalCount;
          perElevCost = cost / totalCount;
          perElevDisc = discounted / totalCount;
        }
        catOrigPerElev += perElevCost;
        catDiscPerElev += perElevDisc;

        const row: string[] = [
          item.description,
          item.part_number || '',
          qty.toFixed(2),
        ];
        if (showPerElev) row.push(perElevQty.toFixed(2));
        row.push(fmtCurrency(cost));
        if (showPerElev) row.push(fmtCurrency(perElevCost));
        row.push(fmtCurrency(discounted));
        if (showPerElev) row.push(fmtCurrency(perElevDisc));
        rows.push(row);
      }

      elevCatTotals[catKey] = catDiscTotal;
      elevCatPerElevTotals[catKey] = catDiscPerElev;

      // Total row
      const singularLabel = SINGULAR_MAP[catKey] ?? catTitle;
      const totalRow: string[] = [`Total ${singularLabel} Cost`];
      for (let i = 1; i < headers.length; i++) totalRow.push('');
      // Fill cost columns in total row
      const costStartIdx = showPerElev ? 4 : 3;
      totalRow[costStartIdx] = fmtCurrency(catOrigTotal);
      if (showPerElev) totalRow[costStartIdx + 1] = fmtCurrency(catOrigPerElev);
      totalRow[showPerElev ? costStartIdx + 2 : costStartIdx + 1] = fmtCurrency(catDiscTotal);
      if (showPerElev) totalRow[costStartIdx + 3] = fmtCurrency(catDiscPerElev);
      rows.push(totalRow);

      // Section title
      doc.setFontSize(9);
      doc.setFont('helvetica', 'bold');
      doc.text(catTitle, margin, currentY);
      currentY += 2;

      // Column styles — right-align numeric columns
      const colStyles: Record<number, { halign?: 'left' | 'center' | 'right'; cellWidth?: number }> = {
        0: { cellWidth: showPerElev ? 50 : 70 },
      };
      for (let c = 2; c < headers.length; c++) {
        colStyles[c] = { halign: 'right' as const };
      }

      autoTable(doc, {
        startY: currentY,
        margin: { left: margin, right: margin },
        head: [headers],
        body: rows,
        styles: { fontSize: 6, cellPadding: 1.2 },
        headStyles: { fillColor: [47, 84, 150], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 5.5 },
        columnStyles: colStyles,
        didParseCell: (data: unknown) => {
          const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
          if (d.section === 'body' && d.row.index === rows.length - 1) {
            d.cell.styles.fontStyle = 'bold';
            d.cell.styles.fillColor = [230, 235, 245];
          }
        },
      });
      currentY = getLastY() + 4;

      // Add new page if running low on space
      if (currentY > pageH - 40) {
        doc.addPage();
        currentY = 20;
      }
    }

    // ---- Elevation Cost Summary ----
    const costSumHeaders: string[] = ['COST/ELEVATION'];
    if (showPerElev) costSumHeaders.push('COST/ELEVATION');
    costSumHeaders.push('TOTAL ELEVATION COST');

    const costSumRows: string[][] = [];
    const costRowDefs: [string, string][] = [
      ['PROFILE COSTS', 'profiles'], ['ACCESSORY COSTS', 'accessories'],
      ['GASKET COSTS', 'gaskets'], ['DOOR COSTS', 'doors'],
      ['GLASS COSTS', 'glass'], ['FABRICATION COSTS', 'fabrication'],
    ];

    let elevTotalCost = 0;
    let elevTotalPerElev = 0;
    for (const [label, key] of costRowDefs) {
      // Skip categories whose material section was unchecked
      if (elevSections?.[key] === false) continue;
      const total = elevCatTotals[key] ?? 0;
      if (total === 0) continue;
      elevTotalCost += total;
      const perElev = elevCatPerElevTotals[key] ?? total / totalCount;
      elevTotalPerElev += perElev;
      const row: string[] = [label];
      if (showPerElev) row.push(fmtCurrency(perElev));
      row.push(fmtCurrency(total));
      costSumRows.push(row);
    }

    // Total row
    const costTotalRow: string[] = [`${elevName} TOTAL COSTS`];
    if (showPerElev) costTotalRow.push(fmtCurrency(elevTotalPerElev));
    costTotalRow.push(fmtCurrency(elevTotalCost));
    costSumRows.push(costTotalRow);

    if (currentY > pageH - 50) {
      doc.addPage();
      currentY = 20;
    }

    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.text('Elevation Cost Summary', margin, currentY);
    currentY += 2;

    autoTable(doc, {
      startY: currentY,
      margin: { left: margin, right: margin },
      head: [costSumHeaders],
      body: costSumRows,
      styles: { fontSize: 7, cellPadding: 1.5 },
      headStyles: { fillColor: [47, 84, 150], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: showPerElev
        ? { 1: { halign: 'right' }, 2: { halign: 'right' } }
        : { 1: { halign: 'right' } },
      didParseCell: (data: unknown) => {
        const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
        if (d.section === 'body' && d.row.index === costSumRows.length - 1) {
          d.cell.styles.fontStyle = 'bold';
          d.cell.styles.fillColor = [32, 55, 100];
          d.cell.styles.textColor = [255, 255, 255];
        }
      },
    });

    // Note
    const noteY = getLastY() + 3;
    doc.setFontSize(6);
    doc.setFont('helvetica', 'italic');
    doc.text('*Note - Elevation costs based on discounted material costs', margin, noteY);
  }

  // ---- Summary Page ----
  doc.addPage();
  doc.setFontSize(16);
  doc.setFont('helvetica', 'bold');
  doc.text('Project Summary', margin, 20);

  // Cost overview — track per-category discounted totals for markup bases
  // Re-price standard parts from scratch for accuracy (matches Excel approach)
  let totalDiscountable = 0;
  let totalNonDiscountable = 0;
  const catDiscounted: Record<string, number> = {
    profiles: 0, accessories: 0, gaskets: 0, doors: 0, glass: 0, fabrication: 0,
  };

  // Fresh materials state for summary re-pricing
  const summaryMaterialsState: Record<string, ExtraMaterial> = {};

  for (const elev of Object.values(elevations)) {
    if (!elev.calculated_outputs) continue;
    const sumFinish = elev.finish || '';
    const summaryOutputs = upgradeGlassOutputs(elev.calculated_outputs, elev, settings.glass_per_sqft ?? 10.5);
    for (const output of summaryOutputs) {
      const cat = classifyOutput(output);
      if (cat === 'calculations') continue;

      let cost: number;
      if (output.manual || cat === 'glass' || cat === 'fabrication' || cat === 'doors') {
        // Use live settings rates for glass/fab
        const qty = sumQty(output.quantity);
        if (cat === 'glass') {
          const glassRate = settings.glass_per_sqft ?? 10.5;
          cost = output.area_sqft != null ? qty * output.area_sqft * glassRate : qty * glassRate;
        } else if (cat === 'fabrication') {
          cost = qty * (settings.fabrication_cost_per_joint ?? 15.0);
        } else {
          cost = output.price ?? 0;
        }
      } else {
        const isGasket = cat === 'gaskets';
        const isProfile = cat === 'profiles';
        const useGroup = isProfile || isGasket;
        const shouldGroup = useGroup && Array.isArray(output.quantity) && output.quantity.length > 1;

        if (shouldGroup) {
          const [price, , impact] = getPriceByPart(
            output.part_number, output.quantity, sumFinish,
            summaryMaterialsState, false, useGroup, output.description,
          );
          if (impact) applyMaterialImpactInMemory(summaryMaterialsState, impact);
          cost = price ?? 0;
        } else {
          let itemTotal = 0;
          const quantities = Array.isArray(output.quantity) ? output.quantity : [output.quantity];
          for (const singleQty of quantities) {
            const [price, , impact] = getPriceByPart(
              output.part_number, singleQty, sumFinish,
              summaryMaterialsState, false, useGroup, output.description,
            );
            if (impact) applyMaterialImpactInMemory(summaryMaterialsState, impact);
            itemTotal += price ?? 0;
          }
          cost = itemTotal;
        }
      }

      if (DISCOUNTABLE_TYPES.has(cat)) {
        totalDiscountable += cost;
        catDiscounted[cat] = (catDiscounted[cat] ?? 0) + cost * multiplier;
      } else {
        totalNonDiscountable += cost;
        catDiscounted[cat] = (catDiscounted[cat] ?? 0) + cost;
      }
    }
  }

  const totalListPrice = totalDiscountable + totalNonDiscountable;
  const discountedTotal = (totalDiscountable * multiplier) + totalNonDiscountable;

  // Waste cost
  let wasteCost = 0;
  for (const [key, mat] of Object.entries(materials)) {
    if (!mat.length_pieces || mat.length_pieces.length === 0) continue;
    const { partNumber, finish } = parseKey(key);
    const [unitPrice] = getUnitPriceByPart(partNumber, finish);
    if (unitPrice != null) {
      wasteCost += mat.length_pieces.reduce((s, l) => s + l, 0) * unitPrice;
    }
  }
  const residualCost = wasteCost * multiplier;
  const wastePct = discountedTotal > 0 ? (residualCost / discountedTotal) * 100 : 0;

  // ---- Elevation Summary Table ----
  const elevSummaryRows = Object.entries(elevations)
    .filter(([, e]) => e.calculated_outputs && e.calculated_outputs.length > 0)
    .map(([name, e]) => {
      const tc = e.total_count || 1;
      const w = e.opening_width_inches ?? 0;
      const h = e.opening_height_inches ?? 0;
      const sqft = (w * h * tc) / 144;
      const perim = (2 * (w + h) * tc) / 12;
      return [name, String(tc), `${w}" x ${h}"`, sqft.toFixed(2), perim.toFixed(2)];
    });

  if (elevSummaryRows.length > 0) {
    // Totals
    let totalQty = 0, totalSqft = 0, totalPerim = 0;
    for (const r of elevSummaryRows) {
      totalQty += Number(r[1]);
      totalSqft += Number(r[3]);
      totalPerim += Number(r[4]);
    }
    elevSummaryRows.push(['TOTAL', String(totalQty), '', totalSqft.toFixed(2), totalPerim.toFixed(2)]);

    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text('ELEVATION SUMMARY', margin, 28);

    autoTable(doc, {
      startY: 31,
      margin: { left: margin, right: margin },
      head: [['Elevation Name', 'Quantity (EA)', 'Dimensions', 'SQFT Total (SQFT)', 'Perimeter FT Total (FT)']],
      body: elevSummaryRows,
      styles: { fontSize: 8, cellPadding: 2 },
      headStyles: { fillColor: [84, 130, 53], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: { 1: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' } },
      didParseCell: (data: unknown) => {
        const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
        if (d.section === 'body' && d.row.index === elevSummaryRows.length - 1) {
          d.cell.styles.fontStyle = 'bold';
        }
      },
    });
  }

  // ---- Cost Overview ----
  const costOverviewData = [
    ['List Price Total:', fmtCurrency(totalListPrice)],
    ['Discounted Total:', fmtCurrency(discountedTotal)],
    ['Residual/Waste Cost:', fmtCurrency(residualCost)],
    ['Waste Percentage:', `${wastePct.toFixed(2)}%`],
  ];

  autoTable(doc, {
    startY: getLastY() + 8,
    margin: { left: margin, right: margin },
    head: [['COST OVERVIEW', '']],
    body: costOverviewData,
    styles: { fontSize: 9, cellPadding: 2.5 },
    headStyles: { fillColor: [47, 84, 150], textColor: [255, 255, 255], fontStyle: 'bold' },
    columnStyles: { 0: { cellWidth: 80 }, 1: { halign: 'right' } },
    didParseCell: (data: unknown) => {
      const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
      if (d.section === 'body' && d.row.index === 1) {
        d.cell.styles.fontStyle = 'bold';
      }
    },
  });

  // ---- Additional Costs ----
  // Category-specific bases:
  //   Overhead Materials → material total only
  //   Overhead Labor → fabrication/labor total only
  //   Rest → project discounted total
  const addMaterialBase = (catDiscounted.profiles ?? 0) + (catDiscounted.accessories ?? 0) +
    (catDiscounted.gaskets ?? 0) + (catDiscounted.doors ?? 0);
  const addLaborBase = catDiscounted.fabrication ?? 0;

  const additionalDefs: [string, number, number][] = [
    ['Overhead Materials', settings.overhead_materials_pct ?? 0, addMaterialBase],
    ['Overhead Labor', settings.overhead_labor_pct ?? 0, addLaborBase],
    ['Admin and Management', settings.admin_management_pct ?? 0, discountedTotal],
    ['Engineering', settings.engineering_pct ?? 0, discountedTotal],
    ['Packaging Materials', settings.packaging_materials_pct ?? 0, discountedTotal],
    ['Shipping and Transport', settings.shipping_transport_pct ?? 0, discountedTotal],
    ['Commissions', settings.commissions_pct ?? 0, discountedTotal],
  ];
  const activeAdditional = additionalDefs.filter(([, pct]) => pct > 0);
  const additionalTotal = activeAdditional.reduce((s, [, pct, base]) => s + base * (pct / 100), 0);

  if (activeAdditional.length > 0) {
    const addRows = activeAdditional.map(([label, pct, base]) => [`${label} (${pct}%)`, fmtCurrency(base * (pct / 100))]);
    addRows.push(['SUBTOTAL', fmtCurrency(additionalTotal)]);

    autoTable(doc, {
      startY: getLastY() + 6,
      margin: { left: margin, right: margin },
      head: [['ADDITIONAL COSTS', '']],
      body: addRows,
      styles: { fontSize: 9, cellPadding: 2.5 },
      headStyles: { fillColor: [84, 130, 53], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: { 0: { cellWidth: 80 }, 1: { halign: 'right' } },
      didParseCell: (data: unknown) => {
        const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
        if (d.section === 'body' && d.row.index === addRows.length - 1) {
          d.cell.styles.fontStyle = 'bold';
        }
      },
    });
  }

  // ---- Markups ----
  const materialBase = (catDiscounted.profiles ?? 0) + (catDiscounted.accessories ?? 0) +
    (catDiscounted.gaskets ?? 0) + (catDiscounted.doors ?? 0);
  const glassBase = catDiscounted.glass ?? 0;
  const laborBase = catDiscounted.fabrication ?? 0;

  const markupDefs: [string, number, number][] = [
    ['Profit on Material', settings.profit_on_material_pct ?? 0, materialBase],
    ['Profit on Waste', settings.profit_on_waste_pct ?? 0, residualCost],
    ['Profit on Glass Purchase', settings.profit_on_glass_pct ?? 0, glassBase],
    ['Profit on Wages', settings.profit_on_wages_pct ?? 0, laborBase],
    ['Planning / Technical Office', settings.planning_technical_pct ?? 0, discountedTotal],
    ['Commission', settings.commission_pct ?? 0, discountedTotal],
  ];
  const activeMarkups = markupDefs.filter(([, pct]) => pct > 0);
  const markupTotal = activeMarkups.reduce((sum, [, pct, base]) => sum + base * (pct / 100), 0);

  if (activeMarkups.length > 0) {
    const mkRows = activeMarkups.map(([label, pct, base]) => [`${label} (${pct}%)`, fmtCurrency(base * (pct / 100))]);
    mkRows.push(['SUBTOTAL', fmtCurrency(markupTotal)]);

    autoTable(doc, {
      startY: getLastY() + 6,
      margin: { left: margin, right: margin },
      head: [['MARKUPS / PROFIT', '']],
      body: mkRows,
      styles: { fontSize: 9, cellPadding: 2.5 },
      headStyles: { fillColor: [112, 48, 160], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: { 0: { cellWidth: 80 }, 1: { halign: 'right' } },
      didParseCell: (data: unknown) => {
        const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
        if (d.section === 'body' && d.row.index === mkRows.length - 1) {
          d.cell.styles.fontStyle = 'bold';
        }
      },
    });
  }

  // ---- Field Costs ----
  let fieldInstallation = 0;
  let fieldSealants = 0;
  let fieldBreakMetal = 0;
  {
    const laborRate = settings.installation_labor_rate ?? 65;
    const laborMkp = 1 + (settings.installation_labor_markup_pct ?? 0) / 100;
    const sealRate = settings.sealant_rate_per_ft ?? 3.5;
    const sealMkp = 1 + (settings.sealant_markup_pct ?? 0) / 100;
    const bmRate = settings.break_metal_rate_per_ft ?? 12;
    const bmMkp = 1 + (settings.break_metal_markup_pct ?? 0) / 100;

    for (const [, elev] of Object.entries(elevations)) {
      if (!elev.calculated_outputs || elev.calculated_outputs.length === 0) continue;
      const qty = elev.total_count || 1;
      const w = elev.opening_width_inches || 0;
      const h = elev.opening_height_inches || 0;
      const perimFt = (2 * (w + h)) / 12;
      const wFt = w / 12;
      const hFt = h / 12;

      if (elev.installation_labor_hours && elev.installation_labor_hours > 0) {
        fieldInstallation += elev.installation_labor_hours * laborRate * qty * laborMkp;
      }
      if (elev.sealant_joints && elev.sealant_joints > 0) {
        fieldSealants += elev.sealant_joints * sealRate * perimFt * qty * sealMkp;
      }
      if (elev.break_metal_selections && elev.break_metal_selections.length > 0) {
        let linFt = 0;
        for (const sel of elev.break_metal_selections) {
          if (sel === 'Perimeter') linFt += 2 * (wFt + hFt);
          else if (sel === 'Head') linFt += wFt;
          else if (sel === 'Sill') linFt += wFt;
          else if (sel === 'Left Jamb') linFt += hFt;
          else if (sel === 'Right Jamb') linFt += hFt;
          else if (sel === 'Both Jambs') linFt += 2 * hFt;
        }
        fieldBreakMetal += linFt * bmRate * qty * bmMkp;
      }
    }
  }

  const liftAmt = settings.lift_equipment_amount ?? 0;
  const liftType = settings.lift_equipment_type ?? 'lump_sum';
  const liftMkp = 1 + (settings.lift_equipment_markup_pct ?? 0) / 100;
  const fieldSubBeforeLift = fieldInstallation + fieldSealants + fieldBreakMetal;
  const subtotalBeforeLift = discountedTotal + additionalTotal + markupTotal + fieldSubBeforeLift;
  const fieldLift = liftType === 'percentage'
    ? subtotalBeforeLift * (liftAmt / 100) * liftMkp
    : liftAmt * liftMkp;
  const fieldCostTotal = fieldSubBeforeLift + fieldLift;

  // Render field costs breakdown table
  if (fieldCostTotal > 0) {
    const fcItems: [string, number][] = [
      ['Installation Labor', fieldInstallation],
      ['Perimeter Sealants', fieldSealants],
      ['Aluminum Break Metal', fieldBreakMetal],
      ['Lift Equipment', fieldLift],
    ];
    const activeFc = fcItems.filter(([, amt]) => amt > 0);
    const fcRows = activeFc.map(([label, amt]) => [label, fmtCurrency(amt)]);
    fcRows.push(['SUBTOTAL', fmtCurrency(fieldCostTotal)]);

    autoTable(doc, {
      startY: getLastY() + 6,
      margin: { left: margin, right: margin },
      head: [['FIELD COSTS & INSTALLATION', '']],
      body: fcRows,
      styles: { fontSize: 9, cellPadding: 2.5 },
      headStyles: { fillColor: [191, 143, 0], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: { 0: { cellWidth: 80 }, 1: { halign: 'right' } },
      didParseCell: (data: unknown) => {
        const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
        if (d.section === 'body' && d.row.index === fcRows.length - 1) {
          d.cell.styles.fontStyle = 'bold';
        }
      },
    });
  }

  // ---- Project Total ----
  const grandTotal = discountedTotal + additionalTotal + markupTotal + fieldCostTotal;
  const ptRows: string[][] = [
    ['Discounted Total:', fmtCurrency(discountedTotal)],
  ];
  if (additionalTotal > 0) ptRows.push(['+ Additional:', fmtCurrency(additionalTotal)]);
  if (markupTotal > 0) ptRows.push(['+ Markups:', fmtCurrency(markupTotal)]);
  if (fieldCostTotal > 0) ptRows.push(['+ Field Costs:', fmtCurrency(fieldCostTotal)]);
  ptRows.push(['GRAND TOTAL:', fmtCurrency(grandTotal)]);

  autoTable(doc, {
    startY: getLastY() + 6,
    margin: { left: margin, right: margin },
    head: [['PROJECT TOTAL', '']],
    body: ptRows,
    styles: { fontSize: 10, cellPadding: 3 },
    headStyles: { fillColor: [47, 84, 150], textColor: [255, 255, 255], fontStyle: 'bold' },
    columnStyles: { 0: { cellWidth: 80 }, 1: { halign: 'right' } },
    didParseCell: (data: unknown) => {
      const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
      if (d.section === 'body' && d.row.index === ptRows.length - 1) {
        d.cell.styles.fontStyle = 'bold';
        d.cell.styles.fillColor = [32, 55, 100];
        d.cell.styles.textColor = [255, 255, 255];
      }
    },
  });

  // ---- Pie Chart ----
  try {
    const activeMaterialCost = Math.max(0, discountedTotal - residualCost);
    const pieBase64 = createPdfPieChart(activeMaterialCost, additionalTotal, markupTotal, residualCost, fieldCostTotal);
    if (pieBase64) {
      const lastYPie = getLastY();
      const chartH = 70;
      const chartW = 80;
      if (lastYPie + chartH + 10 > pageH - margin) {
        doc.addPage();
        doc.addImage(pieBase64, 'PNG', pageW / 2 - chartW / 2, 20, chartW, chartH);
      } else {
        doc.addImage(pieBase64, 'PNG', pageW / 2 - chartW / 2, lastYPie + 8, chartW, chartH);
      }
    }
  } catch (e) {
    console.warn('Could not add pie chart to PDF:', e);
  }

  // ---- Download ----
  doc.save(`${projectName.replace(/[^a-zA-Z0-9_-]/g, '_')}_Report.pdf`);
}
