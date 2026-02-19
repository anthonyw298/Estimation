import type {
  ElevationData,
  ProjectSettings,
  ExtraMaterial,
  CalculatedOutput,
  DoorConfig,
} from '@/types';
import { getUnitPriceByPart } from '@/lib/pricing';

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
): string | null {
  const grandTotal = materialCost + miscCost + markupCost + residualCost;
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

  const MATERIAL_COLOR = '#4472C4';
  const MISC_COLOR = '#548235';
  const MARKUP_COLOR = '#7030A0';
  const RESIDUAL_COLOR = '#ED7D31';

  interface Seg { name: string; value: number; pct: number; color: string }
  const segments: Seg[] = [];
  if (materialCost > 0) segments.push({ name: 'Active Materials', value: materialCost, pct: materialPct, color: MATERIAL_COLOR });
  if (miscCost > 0) segments.push({ name: 'Additional', value: miscCost, pct: miscPct, color: MISC_COLOR });
  if (markupCost > 0) segments.push({ name: 'Profit/Markups', value: markupCost, pct: markupPct, color: MARKUP_COLOR });
  if (residualCost > 0) segments.push({ name: 'Residual/Waste', value: residualCost, pct: residualPct, color: RESIDUAL_COLOR });

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
): Promise<void> {
  // Dynamic import for browser
  const { default: jsPDF } = await import('jspdf');
  const autoTable = (await import('jspdf-autotable')).default;

  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'letter' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 15;

  // Pre-compute multiplier
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

  // ---- Per-Elevation Pages ----
  const sortedNames = Object.keys(elevations).sort();

  for (const elevName of sortedNames) {
    const elev = elevations[elevName];
    if (!elev.calculated_outputs || elev.calculated_outputs.length === 0) continue;

    doc.addPage();
    const elevDoors = doors[elevName] || [];

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
      `Count: ${elev.total_count || 1}`,
      `Doors: ${elevDoors.length > 0 ? elevDoors.map(d => `${d.count}x ${d.size}`).join(', ') : 'None'}`,
    ];
    doc.text(info.join('  |  '), margin, 28);

    // Material table by category
    const catOrder: [string, string][] = [
      ['profiles', 'Profiles'], ['accessories', 'Accessories'], ['gaskets', 'Gaskets'],
      ['doors', 'Doors'], ['glass', 'Glass'], ['fabrication', 'Fabrication'],
    ];

    const tableData: Array<[string, string, string, string, string, string]> = [];
    let elevTotal = 0;

    for (const [catKey, catLabel] of catOrder) {
      const items = elev.calculated_outputs.filter(o => classifyOutput(o) === catKey);
      if (items.length === 0) continue;

      // Category header row
      tableData.push([catLabel.toUpperCase(), '', '', '', '', '']);

      for (const item of items) {
        const qty = sumQty(item.quantity);
        const cost = item.price ?? 0;
        const isDisc = DISCOUNTABLE_TYPES.has(catKey);
        const discounted = isDisc ? cost * multiplier : cost;
        elevTotal += discounted;

        tableData.push([
          item.description,
          item.part_number,
          qty.toFixed(2),
          fmtCurrency(cost),
          isDisc ? `x${multiplier.toFixed(3)}` : '—',
          fmtCurrency(discounted),
        ]);
      }
    }

    // Total row
    tableData.push(['TOTAL', '', '', '', '', fmtCurrency(elevTotal)]);

    autoTable(doc, {
      startY: 33,
      margin: { left: margin, right: margin },
      head: [['Description', 'Part Number', 'Total Quantity Required', 'Total List Cost', 'Multiplier', 'Discounted Total List Cost']],
      body: tableData,
      styles: { fontSize: 7, cellPadding: 1.5 },
      headStyles: { fillColor: [47, 84, 150], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: {
        0: { cellWidth: 80 },
        2: { halign: 'right' },
        3: { halign: 'right' },
        4: { halign: 'center' },
        5: { halign: 'right' },
      },
      didParseCell: (data: unknown) => {
        const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
        if (d.section === 'body') {
          const rowData = tableData[d.row.index];
          if (rowData && rowData[1] === '' && rowData[2] === '' && rowData[3] === '' && d.row.index < tableData.length - 1) {
            d.cell.styles.fontStyle = 'bold';
            d.cell.styles.fillColor = [230, 235, 245];
          }
          if (d.row.index === tableData.length - 1) {
            d.cell.styles.fontStyle = 'bold';
            d.cell.styles.fillColor = [32, 55, 100];
            d.cell.styles.textColor = [255, 255, 255];
          }
        }
      },
    });
  }

  // ---- Summary Page ----
  doc.addPage();
  doc.setFontSize(16);
  doc.setFont('helvetica', 'bold');
  doc.text('Project Summary', margin, 20);

  // Cost overview
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

  // Additional costs
  const additionalPcts = [
    settings.overhead_materials_pct ?? 0, settings.overhead_labor_pct ?? 0,
    settings.admin_management_pct ?? 0, settings.engineering_pct ?? 0,
    settings.packaging_materials_pct ?? 0, settings.shipping_transport_pct ?? 0,
    settings.commissions_pct ?? 0,
  ];
  const additionalTotal = discountedTotal * (additionalPcts.reduce((s, v) => s + v, 0) / 100);

  // Markups
  const markupPcts = [
    settings.profit_on_material_pct ?? 0, settings.profit_on_waste_pct ?? 0,
    settings.profit_on_glass_pct ?? 0, settings.profit_on_wages_pct ?? 0,
    settings.planning_technical_pct ?? 0, settings.commission_pct ?? 0,
  ];
  const markupTotal = discountedTotal * (markupPcts.reduce((s, v) => s + v, 0) / 100);

  const grandTotal = discountedTotal + residualCost + additionalTotal + markupTotal;

  // Summary table
  const summaryData = [
    ['List Price Total', fmtCurrency(totalListPrice)],
    ['Discount Multiplier', `x ${multiplier.toFixed(3)}`],
    ['Discounted Total', fmtCurrency(discountedTotal)],
    ['Residual/Waste Cost', fmtCurrency(residualCost)],
    ['Waste Percentage', `${wastePct.toFixed(2)}%`],
    ...(additionalTotal > 0 ? [['Additional Costs', fmtCurrency(additionalTotal)]] : []),
    ...(markupTotal > 0 ? [['Markups', fmtCurrency(markupTotal)]] : []),
    ['GRAND TOTAL', fmtCurrency(grandTotal)],
  ];

  autoTable(doc, {
    startY: 28,
    margin: { left: margin, right: margin },
    head: [['Item', 'Value']],
    body: summaryData,
    styles: { fontSize: 10, cellPadding: 3 },
    headStyles: { fillColor: [47, 84, 150], textColor: [255, 255, 255], fontStyle: 'bold' },
    columnStyles: {
      0: { cellWidth: 80 },
      1: { halign: 'right' },
    },
    didParseCell: (data: unknown) => {
      const d = data as { row: { index: number }; section: string; cell: { styles: Record<string, unknown> } };
      if (d.section === 'body' && d.row.index === summaryData.length - 1) {
        d.cell.styles.fontStyle = 'bold';
        d.cell.styles.fillColor = [32, 55, 100];
        d.cell.styles.textColor = [255, 255, 255];
      }
    },
  });

  // ---- Pie Chart ----
  try {
    const activeMaterialCost = Math.max(0, discountedTotal - residualCost);
    const pieBase64 = createPdfPieChart(activeMaterialCost, additionalTotal, markupTotal, residualCost);
    if (pieBase64) {
      const lastYPie = (doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? 100;
      // Check if enough space on page, otherwise add new page
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

  // Per-elevation cost breakdown
  const elevCostData = Object.entries(elevations)
    .filter(([, e]) => e.calculated_outputs && e.calculated_outputs.length > 0)
    .map(([name, elev]) => {
      let cost = 0;
      for (const output of elev.calculated_outputs!) {
        if (output.type === 'Calculations' || output.price == null) continue;
        const cat = classifyOutput(output);
        cost += DISCOUNTABLE_TYPES.has(cat) ? output.price * multiplier : output.price;
      }
      return [name, fmtCurrency(cost)];
    });

  if (elevCostData.length > 0) {
    const lastY = (doc as unknown as { lastAutoTable?: { finalY?: number } }).lastAutoTable?.finalY ?? 100;
    autoTable(doc, {
      startY: lastY + 10,
      margin: { left: margin, right: margin },
      head: [['Elevation Name', 'Discounted Total List Cost']],
      body: elevCostData,
      styles: { fontSize: 9, cellPadding: 2.5 },
      headStyles: { fillColor: [84, 130, 53], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: { 1: { halign: 'right' } },
    });
  }

  // ---- Download ----
  doc.save(`${projectName.replace(/[^a-zA-Z0-9_-]/g, '_')}_Report.pdf`);
}
