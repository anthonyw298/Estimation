import type {
  ElevationData,
  ProjectSettings,
  ExtraMaterial,
  CalculatedOutput,
  DoorConfig,
  ReportConfig,
} from '@/types';
import { getUnitPriceByPart, getPriceByPart, applyMaterialImpactInMemory, parseLengthToFeet } from '@/lib/pricing';
import { calculateDloWidth, calculateDloHeight, calculateGlassMakeSize, calculate_total_glass, buildDloGrid } from '@/lib/formulas';
import { partsData } from '@/data/parts-data';
import { PART_NUMBER_MAP } from '@/data/part-number';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type Worksheet = import('exceljs').Worksheet;
type Row = import('exceljs').Row;
type Cell = import('exceljs').Cell;
type Workbook = import('exceljs').Workbook;

// ---------------------------------------------------------------------------
// Color constants (matching Python's openpyxl colors)
// ---------------------------------------------------------------------------

const BLUE_HEADER = '4472C4';       // Cost Overview header
const GREEN_HEADER = '548235';      // Additional Costs header
const ORANGE_HEADER = 'C65911';     // Markups header
const DARK_BLUE = '2F5496';         // Project Total header, Elevation Summary
const VERY_DARK_BLUE = '203764';    // Grand Total row
const LIGHT_GRAY = 'D6DCE4';       // Subheaders / totals background
const WHITE = 'FFFFFF';
const BLACK = '000000';

// Discountable types
const DISCOUNTABLE_TYPES = new Set(['profiles', 'gaskets', 'accessories']);

// Category classification
const GASKET_PART_NUMBERS = new Set(['E2-0052', 'E2-0053', 'E2-0065']);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function sumQty(qty: number | number[]): number {
  return Array.isArray(qty) ? qty.reduce((s, v) => s + Number(v), 0) : Number(qty);
}

function classifyOutput(output: CalculatedOutput): string {
  const pn = output.part_number || '';
  const desc = (output.description || '').toLowerCase();
  const type = (output.type || '').toLowerCase();

  // Glass
  if (pn === 'GLASS_AREA' || type === 'glass') return 'glass';
  // Fabrication / Labor
  if (pn === 'JOINTS_FAB_LABOR' || type === 'joints_fab_labor' || type === 'fabrication' ||
      desc.includes('joints fabrication') || desc.includes('fabrication labor')) return 'fabrication';
  // Door
  if (type === 'door' || type === 'doors') return 'doors';
  // Calculations (info only)
  if (type === 'calculations') return 'calculations';
  // Gasket
  if (desc.includes('gasket') || GASKET_PART_NUMBERS.has(pn)) return 'gaskets';
  // Accessory
  if (type === 'accessory' || type === 'accessories') return 'accessories';
  // Profile (default for catalog parts)
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
// Excel style helpers
// ---------------------------------------------------------------------------

function thinBorder(): import('exceljs').Border {
  return { style: 'thin', color: { argb: 'FF000000' } };
}

function mediumBorder(): import('exceljs').Border {
  return { style: 'medium', color: { argb: 'FF000000' } };
}

function setFill(cell: Cell, color: string) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${color}` } };
}

function setCurrency(cell: Cell) {
  cell.numFmt = '$#,##0.00';
}

function setPercentFmt(cell: Cell) {
  cell.numFmt = '0.00"%"';
}

function setBoldFont(cell: Cell, size?: number, color?: string) {
  cell.font = { bold: true, size: size ?? 11, color: color ? { argb: `FF${color}` } : undefined };
}

function setFont(cell: Cell, opts: { size?: number; color?: string; bold?: boolean; italic?: boolean }) {
  cell.font = {
    bold: opts.bold,
    italic: opts.italic,
    size: opts.size ?? 11,
    color: opts.color ? { argb: `FF${opts.color}` } : undefined,
  };
}

/**
 * Auto-fit column widths by scanning cell content.
 * Matches the Python _autofit_columns behaviour so descriptions are never
 * truncated.  For currency/number columns the minimum width is 14 to avoid
 * '######' in Excel.
 */
function autofitColumns(sheet: Worksheet, startCol: number, endCol: number, minWidth = 8) {
  for (let c = startCol; c <= endCol; c++) {
    let maxLen = 0;
    let hasNumbers = false;
    sheet.getColumn(c).eachCell({ includeEmpty: false }, (cell) => {
      const val = cell.value;
      if (val != null) {
        maxLen = Math.max(maxLen, String(val).length);
        if (typeof val === 'number' || (cell.numFmt && cell.numFmt.includes('$'))) {
          hasNumbers = true;
        }
      }
    });
    if (hasNumbers) maxLen = Math.max(maxLen, 12);
    const width = Math.max(maxLen + 2, hasNumbers ? 14 : minWidth);
    sheet.getColumn(c).width = width;
  }
}

// ---------------------------------------------------------------------------
// Per-Elevation Sheet Builder
// ---------------------------------------------------------------------------

interface CategoryData {
  items: Array<{
    description: string;
    part_number: string;
    total_qty: number;
    qty_display: string;           // formatted with units, e.g. "8ft x 3", "2.00 pcs"
    qty_per_elev: number;
    qty_per_elev_display: string;  // formatted with units
    total_list_cost: number;
    list_cost_per_elev: number;
    discounted_total_cost: number;
    discounted_per_elev: number;
  }>;
  total_original: number;
  total_discounted: number;
  total_original_per_elev: number;
  total_discounted_per_elev: number;
}

/**
 * Build a display string for quantities with units, matching the Python
 * _write_output_section display_qty_string logic.
 * Profiles: "8ft x 3", "3ft x2, 8ft x1"
 * Accessories: "2.00 pcs"
 * Glass: "150.00 sqft"
 */
function formatQtyDisplay(
  rawQty: number | number[],
  cat: string,
  unit?: string,
): string {
  const isProfile = cat === 'profiles';
  const isGasket = cat === 'gaskets';
  const isAccessory = cat === 'accessories';
  const isGlass = cat === 'glass';
  const displayUnit = (isProfile || isGasket) ? 'ft'
    : isAccessory ? 'pcs'
    : isGlass ? 'sqft'
    : unit || 'pcs';

  if (Array.isArray(rawQty)) {
    if (rawQty.length === 0) return `0 ${displayUnit}`;
    // Check if all values are the same
    const allSame = rawQty.every(v => v === rawQty[0]);
    if (rawQty.length > 1 && allSame) {
      const valStr = (isProfile && rawQty[0] === Math.floor(rawQty[0]))
        ? `${Math.floor(rawQty[0])}${displayUnit}`
        : `${rawQty[0].toFixed(2)} ${displayUnit}`;
      return `${valStr} x ${rawQty.length}`;
    }
    // Group identical values with counts: "3ft x2, 8ft x1"
    const counts = new Map<number, number>();
    for (const v of rawQty) {
      const rounded = Math.round(v * 100) / 100;
      counts.set(rounded, (counts.get(rounded) || 0) + 1);
    }
    const parts: string[] = [];
    for (const [val, count] of [...counts.entries()].sort((a, b) => a[0] - b[0])) {
      const valStr = (isProfile && val === Math.floor(val))
        ? `${Math.floor(val)}${displayUnit}`
        : `${val.toFixed(2)} ${displayUnit}`;
      parts.push(count > 1 ? `${valStr} x${count}` : valStr);
    }
    return parts.join(', ');
  }

  // Scalar quantity
  if (isProfile && rawQty === Math.floor(rawQty)) {
    return `${Math.floor(rawQty)}${displayUnit}`;
  }
  return `${rawQty.toFixed(2)} ${displayUnit}`;
}

function buildElevationCategories(
  outputs: CalculatedOutput[],
  finish: string,
  totalCount: number,
  multiplier: number,
  singleElevOutputs?: CalculatedOutput[],
): Record<string, CategoryData> {
  const categories: Record<string, CategoryData> = {
    profiles: { items: [], total_original: 0, total_discounted: 0, total_original_per_elev: 0, total_discounted_per_elev: 0 },
    accessories: { items: [], total_original: 0, total_discounted: 0, total_original_per_elev: 0, total_discounted_per_elev: 0 },
    gaskets: { items: [], total_original: 0, total_discounted: 0, total_original_per_elev: 0, total_discounted_per_elev: 0 },
    doors: { items: [], total_original: 0, total_discounted: 0, total_original_per_elev: 0, total_discounted_per_elev: 0 },
    glass: { items: [], total_original: 0, total_discounted: 0, total_original_per_elev: 0, total_discounted_per_elev: 0 },
    fabrication: { items: [], total_original: 0, total_discounted: 0, total_original_per_elev: 0, total_discounted_per_elev: 0 },
  };

  // ---------------------------------------------------------------------------
  // Build single-elevation price lookup (count=1, no residual).
  // When available, per-elev values come from an independent count=1 calculation
  // rather than dividing total by count.
  // ---------------------------------------------------------------------------
  const singleElevPriceMap = new Map<string, { price: number; quantity: number | number[] }>();
  if (singleElevOutputs && totalCount > 1) {
    for (const sOut of singleElevOutputs) {
      const cat = classifyOutput(sOut);
      if (cat === 'calculations') continue;
      // Key by category + description + part_number to handle duplicates
      const key = `${cat}|${sOut.description}|${sOut.part_number}`;
      singleElevPriceMap.set(key, { price: sOut.price ?? 0, quantity: sOut.quantity });
    }
  }

  // Per-elevation fresh state for inventory tracking (matches Python _write_output_section)
  const elevMaterialsState: Record<string, ExtraMaterial> = {};

  for (const output of outputs) {
    const cat = classifyOutput(output);
    if (cat === 'calculations') continue;
    if (!categories[cat]) continue;

    const qty = sumQty(output.quantity);
    const isDiscountable = DISCOUNTABLE_TYPES.has(cat);

    let totalCost: number;
    if (output.manual || cat === 'glass' || cat === 'fabrication' || cat === 'doors') {
      // Manual items (glass, fabrication, doors): output.price is already the
      // total cost (ElevationEditor stores qty * unitRate). Use it directly.
      totalCost = output.price ?? 0;
    } else {
      // Standard parts: re-price from scratch using getPriceByPart
      // This matches the Python _write_output_section approach
      const isGasket = cat === 'gaskets';
      const isProfile = cat === 'profiles';
      const useGroup = isProfile || isGasket;
      const shouldGroup = useGroup && Array.isArray(output.quantity) && output.quantity.length > 1;

      if (shouldGroup) {
        // Group processing: pass full list for cut optimization
        const [price, , impact] = getPriceByPart(
          output.part_number, output.quantity, finish,
          elevMaterialsState, false, useGroup, output.description,
        );
        if (impact) applyMaterialImpactInMemory(elevMaterialsState, impact);
        totalCost = price ?? 0;
      } else {
        // Standard processing: iterate individual quantities
        let itemTotal = 0;
        const quantities = Array.isArray(output.quantity) ? output.quantity : [output.quantity];
        for (const singleQty of quantities) {
          const [price, , impact] = getPriceByPart(
            output.part_number, singleQty, finish,
            elevMaterialsState, false, useGroup, output.description,
          );
          if (impact) applyMaterialImpactInMemory(elevMaterialsState, impact);
          itemTotal += price ?? 0;
        }
        totalCost = itemTotal;
      }
    }

    // Apply multiplier for discountable categories (profiles, gaskets, accessories)
    const discountedCost = isDiscountable ? totalCost * multiplier : totalCost;

    // Per-elevation values: use single-elev outputs when available, otherwise divide
    const singleKey = `${cat}|${output.description}|${output.part_number}`;
    const singleData = singleElevPriceMap.get(singleKey);
    let perElev: number;
    let discountedPerElev: number;
    let qtyPerElevVal: number;
    let qtyPerElevDisplay: string;

    if (singleData && totalCount > 1) {
      // True per-elevation cost from independent count=1 calculation
      const singleCost = singleData.price;
      const singleDiscounted = isDiscountable ? singleCost * multiplier : singleCost;
      perElev = singleCost;
      discountedPerElev = singleDiscounted;
      qtyPerElevVal = sumQty(singleData.quantity);
      qtyPerElevDisplay = formatQtyDisplay(singleData.quantity, cat, output.unit);
    } else if (totalCount > 1) {
      perElev = totalCost / totalCount;
      discountedPerElev = discountedCost / totalCount;
      qtyPerElevVal = qty / totalCount;
      if (Array.isArray(output.quantity)) {
        const perElvArr = output.quantity.map(v => v / totalCount);
        qtyPerElevDisplay = formatQtyDisplay(perElvArr, cat, output.unit);
      } else {
        qtyPerElevDisplay = formatQtyDisplay(output.quantity / totalCount, cat, output.unit);
      }
    } else {
      perElev = totalCost;
      discountedPerElev = discountedCost;
      qtyPerElevVal = qty;
      qtyPerElevDisplay = formatQtyDisplay(output.quantity, cat, output.unit);
    }

    // Build display strings with units (matching Python display_qty_string)
    const qtyDisplay = formatQtyDisplay(output.quantity, cat, output.unit);

    categories[cat].items.push({
      description: output.description,
      part_number: output.part_number,
      total_qty: qty,
      qty_display: qtyDisplay,
      qty_per_elev: qtyPerElevVal,
      qty_per_elev_display: qtyPerElevDisplay,
      total_list_cost: totalCost,
      list_cost_per_elev: perElev,
      discounted_total_cost: discountedCost,
      discounted_per_elev: discountedPerElev,
    });

    categories[cat].total_original += totalCost;
    categories[cat].total_discounted += discountedCost;
    categories[cat].total_original_per_elev += perElev;
    categories[cat].total_discounted_per_elev += discountedPerElev;
  }

  return categories;
}

function writeSystemInput(
  sheet: Worksheet,
  elevName: string,
  elev: ElevationData,
  doors: DoorConfig[],
  startRow: number,
): number {
  const rows: [string, string][] = [
    ['System Input', elev.door_only ? 'Door Only' : elev.system_type],
    ['Finish', elev.finish],
    ['Elevation Type', elevName],
    ['Total Count', String(elev.total_count || 1)],
    ['Bays Wide', String(elev.bays_wide || 0)],
    ['Bays Tall', String(elev.bays_tall || 0)],
    ['Custom Bay Widths', elev.custom_bay_widths?.length
      ? elev.custom_bay_widths.map(w => w.toFixed(2) + ' in').join(', ')
      : 'Equal distribution'],
    ['Custom Bay Heights', elev.custom_bay_heights?.length
      ? elev.custom_bay_heights.map(h => h.toFixed(2) + ' in').join(', ')
      : 'Equal distribution'],
    ['Opening Width', `${(elev.opening_width_inches || 0).toFixed(2)} in`],
    ['Opening Height', `${(elev.opening_height_inches || 0).toFixed(2)} in`],
    ['Sq Ft per Type (DLO)', `${calculate_total_glass(elev.opening_width_inches, elev.opening_height_inches, 1, elev.bays_wide || 1, elev.bays_tall || 1, elev.custom_bay_widths, elev.custom_bay_heights).toFixed(2)} sqft`],
    ['Total Sq Ft (DLO)', `${calculate_total_glass(elev.opening_width_inches, elev.opening_height_inches, elev.total_count || 1, elev.bays_wide || 1, elev.bays_tall || 1, elev.custom_bay_widths, elev.custom_bay_heights).toFixed(2)} sqft`],
    ['Perimeter Ft', `${((2 * (elev.opening_width_inches + elev.opening_height_inches)) / 12).toFixed(2)} ft`],
    ['Total Perimeter Ft', `${(((2 * (elev.opening_width_inches + elev.opening_height_inches)) / 12) * (elev.total_count || 1)).toFixed(2)} ft`],
    ['Doors', doors.length > 0
      ? doors.map(d => `${d.count}x ${d.size} (${d.stile})`).join('; ')
      : 'None'],
  ];

  let row = startRow;
  for (const [label, value] of rows) {
    const r = sheet.getRow(row);
    r.getCell(1).value = label;
    r.getCell(2).value = value;
    setBoldFont(r.getCell(1), 10);
    setFont(r.getCell(2), { size: 10 });
    r.getCell(1).border = { top: thinBorder(), bottom: thinBorder(), left: thinBorder(), right: thinBorder() };
    r.getCell(2).border = { top: thinBorder(), bottom: thinBorder(), left: thinBorder(), right: thinBorder() };
    row++;
  }

  return row;
}

function writeMaterialSection(
  sheet: Worksheet,
  title: string,
  catData: CategoryData,
  totalCount: number,
  showPerElev: boolean,
  startRow: number,
  startCol: number,
): number {
  if (catData.items.length === 0) return startRow;

  let row = startRow;

  // Section title
  const titleCell = sheet.getRow(row).getCell(startCol);
  titleCell.value = title.toUpperCase();
  setBoldFont(titleCell, 12);
  row++;

  // Header row
  const headers: string[] = ['Description', 'Part Number', 'Total Quantity Required'];
  if (showPerElev) headers.push('Quantity Per Elevation');
  headers.push('Total List Cost');
  if (showPerElev) headers.push('Total List Cost Per Elevation');
  headers.push('Discounted Total List Cost');
  if (showPerElev) headers.push('Discounted Total List Cost Per Elevation');

  const headerRow = sheet.getRow(row);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(startCol + i);
    cell.value = h;
    setBoldFont(cell, 10);
    cell.border = { bottom: thinBorder() };
  });
  row++;

  // Data rows
  for (const item of catData.items) {
    const dataRow = sheet.getRow(row);
    let col = startCol;
    dataRow.getCell(col++).value = item.description;
    dataRow.getCell(col++).value = item.part_number;
    dataRow.getCell(col).value = item.qty_display;
    col++;
    if (showPerElev) {
      dataRow.getCell(col).value = item.qty_per_elev_display;
      col++;
    }
    dataRow.getCell(col).value = item.total_list_cost;
    setCurrency(dataRow.getCell(col));
    col++;
    if (showPerElev) {
      dataRow.getCell(col).value = item.list_cost_per_elev;
      setCurrency(dataRow.getCell(col));
      col++;
    }
    dataRow.getCell(col).value = item.discounted_total_cost;
    setCurrency(dataRow.getCell(col));
    col++;
    if (showPerElev) {
      dataRow.getCell(col).value = item.discounted_per_elev;
      setCurrency(dataRow.getCell(col));
      col++;
    }
    row++;
  }

  // Totals row — use singular category name matching Excel's title_mapping
  const _titleMap: Record<string, string> = {
    'Profiles': 'Profile', 'Accessories': 'Accessory', 'Gaskets': 'Gasket',
    'Doors': 'Door', 'Glass': 'Glass', 'Labor': 'Labor',
  };
  const totalsRow = sheet.getRow(row);
  let tCol = startCol;
  totalsRow.getCell(tCol).value = `Total ${_titleMap[title] ?? title} Cost`;
  setBoldFont(totalsRow.getCell(tCol), 10);
  tCol++; tCol++; tCol++; // skip part #, qty
  if (showPerElev) tCol++;
  totalsRow.getCell(tCol).value = catData.total_original;
  setCurrency(totalsRow.getCell(tCol));
  setBoldFont(totalsRow.getCell(tCol), 10);
  totalsRow.getCell(tCol).border = { top: thinBorder() };
  tCol++;
  if (showPerElev) tCol++;
  totalsRow.getCell(tCol).value = catData.total_discounted;
  setCurrency(totalsRow.getCell(tCol));
  setBoldFont(totalsRow.getCell(tCol), 10);
  totalsRow.getCell(tCol).border = { top: thinBorder() };
  row += 2; // blank row after

  return row;
}

function writeElevationCostSummary(
  sheet: Worksheet,
  categories: Record<string, CategoryData>,
  elevName: string,
  totalCount: number,
  startRow: number,
  startCol: number,
  includedSections?: Record<string, boolean>,
): number {
  let row = startRow;

  // Header
  const headerRow = sheet.getRow(row);
  headerRow.getCell(startCol).value = 'COST/ELEVATION';
  headerRow.getCell(startCol + 1).value = 'COST/ELEVATION';
  setBoldFont(headerRow.getCell(startCol + 1), 10);
  headerRow.getCell(startCol + 2).value = 'TOTAL ELEVATION COST';
  setBoldFont(headerRow.getCell(startCol + 2), 10);
  [startCol, startCol + 1, startCol + 2].forEach(c => {
    headerRow.getCell(c).border = { bottom: thinBorder() };
  });
  row++;

  const catOrder: [string, string][] = [
    ['profiles', 'PROFILE COSTS'],
    ['accessories', 'ACCESSORY COSTS'],
    ['gaskets', 'GASKET COSTS'],
    ['doors', 'DOOR COSTS'],
    ['glass', 'GLASS COSTS'],
    ['fabrication', 'FABRICATION COSTS'],
  ];

  let totalCostPerElev = 0;
  let totalCostAll = 0;

  for (const [key, label] of catOrder) {
    // Skip categories whose material section was unchecked
    if (includedSections?.[key] === false) continue;
    const cat = categories[key];
    if (cat.total_discounted === 0) continue;

    const r = sheet.getRow(row);
    r.getCell(startCol).value = label;
    // Use pre-computed per-elev totals (from single-elev calculation when available)
    const perElev = cat.total_discounted_per_elev;
    r.getCell(startCol + 1).value = perElev;
    setCurrency(r.getCell(startCol + 1));
    r.getCell(startCol + 2).value = cat.total_discounted;
    setCurrency(r.getCell(startCol + 2));
    totalCostPerElev += perElev;
    totalCostAll += cat.total_discounted;
    row++;
  }

  // Separator
  const sepRow = sheet.getRow(row);
  [startCol, startCol + 1, startCol + 2].forEach(c => {
    sepRow.getCell(c).border = { top: thinBorder() };
  });

  // Total
  const totalRow = sheet.getRow(row);
  totalRow.getCell(startCol).value = `${elevName} TOTAL COSTS`;
  setBoldFont(totalRow.getCell(startCol), 10);
  totalRow.getCell(startCol + 1).value = totalCostPerElev;
  setCurrency(totalRow.getCell(startCol + 1));
  setBoldFont(totalRow.getCell(startCol + 1), 10);
  totalRow.getCell(startCol + 2).value = totalCostAll;
  setCurrency(totalRow.getCell(startCol + 2));
  setBoldFont(totalRow.getCell(startCol + 2), 10);
  row++;

  // Note
  const noteRow = sheet.getRow(row);
  noteRow.getCell(startCol).value = '*Note - Elevation costs based on discounted material costs';
  setFont(noteRow.getCell(startCol), { size: 9, italic: true, color: '808080' });
  row += 2;

  return row;
}

// ---------------------------------------------------------------------------
// Summary Sheet Builder
// ---------------------------------------------------------------------------

interface SummaryItem {
  description: string;
  part_number: string;
  project_total_materials: string; // Python col 1: "BE9-2513 (black)" – part number with finish
  quantity_req_ft: string;         // Python col 2: "Total Feet Required" / "Total Pieces Required" / "N/A"
  qty_stick_req: string;           // Python col 3: "Sticks Required" / "Rolls Required" / "Quantity Per Order" / "Unit Price"
  quantity_display: string;        // Python col 4: "Total Quantity Required" / "Orders Required"
  total_qty_required: number;
  unit_price: number;
  total_list_cost: number;
  discounted_total: number;
  residual_qty: number;
  residual_waste_pct: number;
  residual_cost: number;
  reusable_qty_display: string;    // Python col 8: formatted residual qty (grouped pieces)
  reusable_pct_display: string;    // Python col 9: formatted residual pct
}

interface SummaryCategory {
  items: SummaryItem[];
  total_original: number;
  total_discounted: number;
  total_residual: number;
}

function buildSummaryCategories(
  elevations: Record<string, ElevationData>,
  materials: Record<string, ExtraMaterial>,
  multiplier: number,
): Record<string, SummaryCategory> {
  // Step 2: Aggregate quantities across all elevations by category and part number
  // (matching Python create_summary_sheet Steps 2 + 2.5)
  const partMap = new Map<string, {
    category: string;
    description: string;
    part_number: string;
    total_qty: number;
    quantity_list: number[];
    manual_total_cost: number;  // only for manual/glass/fab/door items
    finish: string;
    isDiscountable: boolean;
    isManual: boolean;
  }>();

  for (const [, elev] of Object.entries(elevations)) {
    if (!elev.calculated_outputs) continue;
    const elevFinish = elev.finish || '';
    for (const output of elev.calculated_outputs) {
      const cat = classifyOutput(output);
      if (cat === 'calculations') continue;
      const isManual = !!(output.manual) || cat === 'glass' || cat === 'fabrication' || cat === 'doors';

      // Key construction matching Python logic
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
        if (isManual) {
          // output.price is already the total cost for this elevation
          existing.manual_total_cost += output.price ?? 0;
        }
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

  // Step 3: Re-price all aggregated items from scratch with fresh state
  // (matching Python create_summary_sheet Step 3)
  const summaryMaterialsState: Record<string, ExtraMaterial> = {};

  const categories: Record<string, SummaryCategory> = {
    profiles: { items: [], total_original: 0, total_discounted: 0, total_residual: 0 },
    accessories: { items: [], total_original: 0, total_discounted: 0, total_residual: 0 },
    gaskets: { items: [], total_original: 0, total_discounted: 0, total_residual: 0 },
    doors: { items: [], total_original: 0, total_discounted: 0, total_residual: 0 },
    glass: { items: [], total_original: 0, total_discounted: 0, total_residual: 0 },
    fabrication: { items: [], total_original: 0, total_discounted: 0, total_residual: 0 },
  };

  for (const [, data] of partMap) {
    const cat = categories[data.category];
    if (!cat) continue;

    const isProfile = data.category === 'profiles';
    const isGasket = data.category === 'gaskets';
    const isAccessory = data.category === 'accessories';

    // Compute total_cost: re-price standard parts; use stored price for manual items
    let totalCost: number;
    if (data.isManual) {
      totalCost = data.manual_total_cost;
    } else {
      const useGroup = isGasket;
      // For profiles/gaskets, use quantity_list for cut optimization
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
        summaryMaterialsState, false, useGroup, data.description,
      );
      if (impact) applyMaterialImpactInMemory(summaryMaterialsState, impact);
      totalCost = price ?? 0;
    }

    const discounted = data.isDiscountable ? totalCost * multiplier : totalCost;
    const [unitPrice] = getUnitPriceByPart(data.part_number, data.finish);

    // Compute residual/waste from the summary materials state
    const extraKey = (isProfile || isGasket) && data.finish
      ? `${data.part_number}-${data.finish.toLowerCase()}`
      : data.part_number;
    const partState = summaryMaterialsState[extraKey];
    let residualQty = 0;
    let residualCost = 0;
    if (partState) {
      if (partState.length_pieces && partState.length_pieces.length > 0) {
        residualQty = partState.length_pieces.reduce((s, l) => s + l, 0);
      } else {
        residualQty = partState.quantity ?? 0;
      }
      if (unitPrice != null && residualQty > 0) {
        residualCost = residualQty * unitPrice * multiplier;
      }
    }

    const totalAcquired = data.total_qty + residualQty;
    const residualPct = totalAcquired > 0 ? Math.min((residualQty / totalAcquired) * 100, 100.0) : 0;

    const numElevations = Object.keys(elevations).length;

    // Compute "Project Total Materials" display (part number with finish)
    const projectTotalMaterials = (isProfile || isGasket) && data.finish
      ? `${data.part_number} (${data.finish})`
      : data.part_number;

    // Compute category-specific column values (matching Python)
    let quantityReqFt = 'N/A';
    let qtyStickReq = 'N/A';
    let quantityDisplay = `${data.total_qty.toFixed(2)}`;

    if ((isProfile || isGasket) && data.part_number && data.part_number !== 'N/A') {
      quantityReqFt = `${data.total_qty.toFixed(2)} ft`;
      const partInfo = partsData[data.part_number];
      const lengthStr = partInfo?.['Length'] ?? '';
      const minPurchaseLength = parseLengthToFeet(lengthStr) || 24.0;
      if (minPurchaseLength > 0) {
        const numUnits = Math.ceil(data.total_qty / minPurchaseLength);
        const unitLabel = isGasket ? 'rolls' : 'sticks';
        qtyStickReq = `${numUnits} (${minPurchaseLength.toFixed(0)}ft per)`;
        // Show grouped individual cut list: "10ft x70, 6ft x10, 5ft x20"
        quantityDisplay = formatQtyDisplay(data.quantity_list, data.category);
      }
    } else if (isAccessory && data.part_number && data.part_number !== 'N/A') {
      const partInfo = partsData[data.part_number];
      const unitsStr = partInfo?.['Units'] ?? '1 pcs.';
      const lengthStr = partInfo?.['Length'] ?? '';
      const lengthFt = lengthStr ? parseLengthToFeet(lengthStr) : 0;
      let unitCountPerBundle = 1;
      let unitLabel = 'pcs per';
      if (lengthFt > 1.0) {
        unitCountPerBundle = lengthFt;
        unitLabel = 'ft per';
      } else {
        const match = unitsStr.toLowerCase().match(/^(\d+)\s*pc/);
        if (match) unitCountPerBundle = parseInt(match[1]) || 1;
      }
      quantityReqFt = `${data.total_qty.toFixed(2)} pcs`;
      qtyStickReq = `${unitCountPerBundle.toFixed(0)} ${unitLabel}`;
      const numOrders = unitCountPerBundle > 0 ? Math.ceil(data.total_qty / unitCountPerBundle) : 0;
      quantityDisplay = `${numOrders} order${numOrders !== 1 ? 's' : ''}`;
    } else {
      // glass, doors, fabrication/labor
      quantityReqFt = 'N/A';
      if (data.total_qty > 0) {
        const up = totalCost / data.total_qty;
        qtyStickReq = `$${up.toFixed(2)}`;
      } else {
        qtyStickReq = '$0.00';
      }
      quantityDisplay = `${data.total_qty.toFixed(2)}`;
    }

    // Format residual display — show grouped individual leftover pieces
    let reusableQtyDisplay = 'N/A';
    if ((isProfile || isGasket) && partState && partState.length_pieces && partState.length_pieces.length > 0) {
      // Group leftover pieces: "4ft x18, 3ft x1, 1.50ft x1"
      reusableQtyDisplay = formatQtyDisplay(partState.length_pieces, data.category);
    } else if (isAccessory && residualQty > 0) {
      reusableQtyDisplay = `${residualQty.toFixed(2)} pcs`;
    }
    const reusablePctDisplay = (isProfile || isGasket || isAccessory) && typeof residualPct === 'number' && residualPct > 0
      ? `${residualPct.toFixed(2)}%`
      : 'N/A';

    cat.items.push({
      description: data.description,
      part_number: data.part_number,
      project_total_materials: projectTotalMaterials,
      quantity_req_ft: quantityReqFt,
      qty_stick_req: qtyStickReq,
      quantity_display: quantityDisplay,
      total_qty_required: data.total_qty,
      unit_price: unitPrice ?? 0,
      total_list_cost: totalCost,
      discounted_total: discounted,
      residual_qty: residualQty,
      residual_waste_pct: residualPct,
      residual_cost: residualCost,
      reusable_qty_display: reusableQtyDisplay,
      reusable_pct_display: reusablePctDisplay,
    });

    cat.total_original += totalCost;
    cat.total_discounted += discounted;
    cat.total_residual += residualCost;
  }

  return categories;
}

// Per-category header definitions matching Python exactly (10 columns each)
const SUMMARY_HEADERS: Record<string, string[]> = {
  profiles: [
    'Description', 'Project Total Materials', 'Total Feet Required', 'Sticks Required',
    'Total Quantity Required', 'Total List Cost', 'Discounted Total List Cost',
    'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost',
  ],
  accessories: [
    'Description', 'Project Total Materials', 'Total Pieces Required', 'Quantity Per Order',
    'Orders Required', 'Total List Cost', 'Discounted Total List Cost',
    'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost',
  ],
  gaskets: [
    'Description', 'Project Total Materials', 'Total Feet Required', 'Rolls Required',
    'Total Quantity Required', 'Total List Cost', 'Discounted Total List Cost',
    'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost',
  ],
  glass: [
    'Description', 'Project Total Materials', 'N/A', 'Unit Price',
    'Total Quantity Required', 'Total List Cost', 'Discounted Total List Cost',
    'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost',
  ],
  doors: [
    'Description', 'Project Total Materials', 'N/A', 'Unit Price',
    'Total Quantity Required', 'Total List Cost', 'Discounted Total List Cost',
    'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost',
  ],
  fabrication: [
    'Description', 'Project Total Materials', 'N/A', 'Unit Price',
    'Total Quantity Required', 'Total List Cost', 'Discounted Total List Cost',
    'Residual Material Quantity', 'Residual Waste %', 'Residual Material Cost',
  ],
};

function writeSummaryCategorySection(
  sheet: Worksheet,
  title: string,
  catKey: string,
  cat: SummaryCategory,
  startRow: number,
  startCol: number,
): number {
  if (cat.items.length === 0) return startRow;

  let row = startRow;

  // Title
  const titleCell = sheet.getRow(row).getCell(startCol);
  titleCell.value = title.toUpperCase();
  setBoldFont(titleCell, 12);
  row++;

  // Per-category headers (10 columns, matching Python)
  const headers = SUMMARY_HEADERS[catKey] ?? SUMMARY_HEADERS.profiles;
  const headerRow = sheet.getRow(row);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(startCol + i);
    cell.value = h;
    setBoldFont(cell, 10);
    cell.border = { bottom: thinBorder() };
  });
  row++;

  // Data rows (10 columns matching Python's _get_item_values mapping)
  for (const item of cat.items) {
    const r = sheet.getRow(row);
    // Col 0: Description
    r.getCell(startCol + 0).value = item.description;
    // Col 1: Project Total Materials (part number with finish)
    r.getCell(startCol + 1).value = item.project_total_materials;
    // Col 2: quantity_req_ft (Total Feet Required / Total Pieces Required / N/A)
    r.getCell(startCol + 2).value = item.quantity_req_ft;
    // Col 3: qty_stick_req (Sticks Required / Rolls Required / Qty Per Order / Unit Price)
    r.getCell(startCol + 3).value = item.qty_stick_req;
    // Col 4: quantity_display (Total Quantity Required / Orders Required)
    r.getCell(startCol + 4).value = item.quantity_display;
    // Col 5: Total List Cost
    r.getCell(startCol + 5).value = item.total_list_cost;
    setCurrency(r.getCell(startCol + 5));
    // Col 6: Discounted Total List Cost
    r.getCell(startCol + 6).value = item.discounted_total;
    setCurrency(r.getCell(startCol + 6));
    // Col 7: Residual Material Quantity
    r.getCell(startCol + 7).value = item.reusable_qty_display;
    // Col 8: Residual Waste %
    r.getCell(startCol + 8).value = item.reusable_pct_display;
    // Col 9: Residual Material Cost
    r.getCell(startCol + 9).value = item.residual_cost;
    setCurrency(r.getCell(startCol + 9));
    row++;
  }

  // Category label mapping (matching Python)
  const categoryMapping: Record<string, string> = {
    profiles: 'Profile', accessories: 'Accessory', gaskets: 'Gasket',
    doors: 'Door', glass: 'Glass', fabrication: 'Labor',
  };
  const categoryLabel = categoryMapping[catKey] ?? title;

  // Totals row
  const totRow = sheet.getRow(row);
  totRow.getCell(startCol).value = `Total ${categoryLabel} Cost`;
  setBoldFont(totRow.getCell(startCol), 10);
  // Total List Cost (col 5)
  totRow.getCell(startCol + 5).value = cat.total_original;
  setCurrency(totRow.getCell(startCol + 5));
  setBoldFont(totRow.getCell(startCol + 5), 10);
  totRow.getCell(startCol + 5).border = { top: thinBorder() };
  // Discounted Total (col 6)
  totRow.getCell(startCol + 6).value = cat.total_discounted;
  setCurrency(totRow.getCell(startCol + 6));
  setBoldFont(totRow.getCell(startCol + 6), 10);
  totRow.getCell(startCol + 6).border = { top: thinBorder() };
  // Residual Cost total (col 9)
  totRow.getCell(startCol + 9).value = cat.total_residual;
  setCurrency(totRow.getCell(startCol + 9));
  setBoldFont(totRow.getCell(startCol + 9), 10);
  totRow.getCell(startCol + 9).border = { top: thinBorder() };
  row += 2;

  return row;
}

function writeCostOverviewBox(
  sheet: Worksheet,
  totalListPrice: number,
  discountedTotal: number,
  residualCost: number,
  wastePct: number,
  startRow: number,
  startCol: number,
): number {
  let row = startRow;

  // Header
  const hr = sheet.getRow(row);
  hr.getCell(startCol).value = 'COST OVERVIEW';
  setFill(hr.getCell(startCol), BLUE_HEADER);
  setFont(hr.getCell(startCol), { bold: true, size: 11, color: WHITE });
  hr.getCell(startCol).border = { top: mediumBorder(), left: mediumBorder(), right: mediumBorder(), bottom: thinBorder() };
  setFill(hr.getCell(startCol + 1), BLUE_HEADER);
  hr.getCell(startCol + 1).border = { top: mediumBorder(), right: mediumBorder(), bottom: thinBorder() };
  setFill(hr.getCell(startCol + 2), BLUE_HEADER);
  hr.getCell(startCol + 2).border = { top: mediumBorder(), right: mediumBorder(), bottom: thinBorder() };
  row++;

  const writeBoxRow = (label: string, value: number | string, isLast: boolean, bold?: boolean) => {
    const r = sheet.getRow(row);
    r.getCell(startCol).value = label;
    setFont(r.getCell(startCol), { bold: bold ?? false, size: 10 });
    r.getCell(startCol).border = { left: mediumBorder(), bottom: isLast ? mediumBorder() : undefined };
    r.getCell(startCol + 2).value = typeof value === 'number' ? value : value;
    if (typeof value === 'number') {
      setCurrency(r.getCell(startCol + 2));
    }
    setFont(r.getCell(startCol + 2), { bold: bold ?? false, size: 10 });
    r.getCell(startCol + 2).alignment = { horizontal: 'right' };
    r.getCell(startCol + 2).border = { right: mediumBorder(), bottom: isLast ? mediumBorder() : undefined };
    r.getCell(startCol + 1).border = { bottom: isLast ? mediumBorder() : undefined };
    row++;
  };

  writeBoxRow('List Price Total:', totalListPrice, false);
  writeBoxRow('Discounted Total:', discountedTotal, false, true);
  writeBoxRow('Residual/Waste Cost:', residualCost, false);
  writeBoxRow('Waste Percentage:', `${wastePct.toFixed(2)}%`, true);

  return row;
}

function writeAdditionalCostsSection(
  sheet: Worksheet,
  settings: ProjectSettings,
  baseAmount: number,
  startRow: number,
  startCol: number,
): { row: number; total: number } {
  let row = startRow;

  // Header
  const hr = sheet.getRow(row);
  hr.getCell(startCol).value = 'ADDITIONAL COSTS';
  setFill(hr.getCell(startCol), GREEN_HEADER);
  setFont(hr.getCell(startCol), { bold: true, size: 11, color: WHITE });
  hr.getCell(startCol).border = { top: mediumBorder(), left: mediumBorder(), bottom: thinBorder() };
  setFill(hr.getCell(startCol + 1), GREEN_HEADER);
  hr.getCell(startCol + 1).border = { top: mediumBorder(), bottom: thinBorder() };
  setFill(hr.getCell(startCol + 2), GREEN_HEADER);
  hr.getCell(startCol + 2).border = { top: mediumBorder(), right: mediumBorder(), bottom: thinBorder() };
  row++;

  const items: [string, number, number][] = [
    ['Overhead Materials', settings.overhead_materials_pct ?? 0, 0],
    ['Overhead Labor', settings.overhead_labor_pct ?? 0, 0],
    ['Admin and Management', settings.admin_management_pct ?? 0, 0],
    ['Engineering', settings.engineering_pct ?? 0, 0],
    ['Packaging Materials', settings.packaging_materials_pct ?? 0, 0],
    ['Shipping and Transport', settings.shipping_transport_pct ?? 0, 0],
    ['Commissions', settings.commissions_pct ?? 0, 0],
  ].map(([label, pct]) => [label as string, pct as number, baseAmount * ((pct as number) / 100)]);

  let total = 0;
  const activeItems = items.filter(([, pct]) => pct > 0);

  if (activeItems.length === 0) {
    const r = sheet.getRow(row);
    r.getCell(startCol).value = '(None configured)';
    setFont(r.getCell(startCol), { italic: true, size: 10, color: '808080' });
    r.getCell(startCol).border = { left: mediumBorder() };
    r.getCell(startCol + 2).border = { right: mediumBorder() };
    row++;
  } else {
    for (const [label, pct, amount] of activeItems) {
      const r = sheet.getRow(row);
      r.getCell(startCol).value = `${label} (${pct}%)`;
      setFont(r.getCell(startCol), { size: 10 });
      r.getCell(startCol).border = { left: mediumBorder() };
      r.getCell(startCol + 2).value = amount;
      setCurrency(r.getCell(startCol + 2));
      r.getCell(startCol + 2).alignment = { horizontal: 'right' };
      r.getCell(startCol + 2).border = { right: mediumBorder() };
      total += amount;
      row++;
    }
  }

  // Subtotal
  const sr = sheet.getRow(row);
  sr.getCell(startCol).value = 'SUBTOTAL';
  setBoldFont(sr.getCell(startCol), 10);
  sr.getCell(startCol).border = { left: mediumBorder(), bottom: mediumBorder(), top: thinBorder() };
  sr.getCell(startCol + 1).border = { bottom: mediumBorder(), top: thinBorder() };
  sr.getCell(startCol + 2).value = total;
  setCurrency(sr.getCell(startCol + 2));
  setBoldFont(sr.getCell(startCol + 2), 10);
  sr.getCell(startCol + 2).alignment = { horizontal: 'right' };
  sr.getCell(startCol + 2).border = { right: mediumBorder(), bottom: mediumBorder(), top: thinBorder() };
  row++;

  return { row, total };
}

function writeMarkupsSection(
  sheet: Worksheet,
  settings: ProjectSettings,
  categories: Record<string, SummaryCategory>,
  discountedTotal: number,
  residualTotal: number,
  startRow: number,
  startCol: number,
): { row: number; total: number } {
  let row = startRow;

  // Header
  const hr = sheet.getRow(row);
  hr.getCell(startCol).value = 'MARKUPS / PROFIT';
  setFill(hr.getCell(startCol), ORANGE_HEADER);
  setFont(hr.getCell(startCol), { bold: true, size: 11, color: WHITE });
  hr.getCell(startCol).border = { top: mediumBorder(), left: mediumBorder(), bottom: thinBorder() };
  setFill(hr.getCell(startCol + 1), ORANGE_HEADER);
  hr.getCell(startCol + 1).border = { top: mediumBorder(), bottom: thinBorder() };
  setFill(hr.getCell(startCol + 2), ORANGE_HEADER);
  hr.getCell(startCol + 2).border = { top: mediumBorder(), right: mediumBorder(), bottom: thinBorder() };
  row++;

  // Material base = profiles + accessories + gaskets + doors (discounted)
  const materialBase = (categories.profiles?.total_discounted ?? 0) +
    (categories.accessories?.total_discounted ?? 0) +
    (categories.gaskets?.total_discounted ?? 0) +
    (categories.doors?.total_discounted ?? 0);
  const glassBase = categories.glass?.total_discounted ?? 0;
  const laborBase = categories.fabrication?.total_discounted ?? 0;

  const markups: [string, number, number][] = [
    ['Profit on Material', settings.profit_on_material_pct ?? 0, materialBase],
    ['Profit on Waste', settings.profit_on_waste_pct ?? 0, residualTotal],
    ['Profit on Glass Purchase', settings.profit_on_glass_pct ?? 0, glassBase],
    ['Profit on Wages', settings.profit_on_wages_pct ?? 0, laborBase],
    ['Planning / Technical Office', settings.planning_technical_pct ?? 0, discountedTotal],
    ['Commission', settings.commission_pct ?? 0, discountedTotal],
  ];

  let total = 0;
  const activeMarkups = markups.filter(([, pct]) => pct > 0);

  if (activeMarkups.length === 0) {
    const r = sheet.getRow(row);
    r.getCell(startCol).value = '(None configured)';
    setFont(r.getCell(startCol), { italic: true, size: 10, color: '808080' });
    r.getCell(startCol).border = { left: mediumBorder() };
    r.getCell(startCol + 2).border = { right: mediumBorder() };
    row++;
  } else {
    for (const [label, pct, base] of activeMarkups) {
      const amount = base * (pct / 100);
      const r = sheet.getRow(row);
      r.getCell(startCol).value = `${label} (${pct}%)`;
      setFont(r.getCell(startCol), { size: 10 });
      r.getCell(startCol).border = { left: mediumBorder() };
      r.getCell(startCol + 2).value = amount;
      setCurrency(r.getCell(startCol + 2));
      r.getCell(startCol + 2).alignment = { horizontal: 'right' };
      r.getCell(startCol + 2).border = { right: mediumBorder() };
      total += amount;
      row++;
    }
  }

  // Subtotal
  const sr = sheet.getRow(row);
  sr.getCell(startCol).value = 'SUBTOTAL';
  setBoldFont(sr.getCell(startCol), 10);
  sr.getCell(startCol).border = { left: mediumBorder(), bottom: mediumBorder(), top: thinBorder() };
  sr.getCell(startCol + 1).border = { bottom: mediumBorder(), top: thinBorder() };
  sr.getCell(startCol + 2).value = total;
  setCurrency(sr.getCell(startCol + 2));
  setBoldFont(sr.getCell(startCol + 2), 10);
  sr.getCell(startCol + 2).alignment = { horizontal: 'right' };
  sr.getCell(startCol + 2).border = { right: mediumBorder(), bottom: mediumBorder(), top: thinBorder() };
  row++;

  return { row, total };
}

function writeProjectTotalSection(
  sheet: Worksheet,
  discountedTotal: number,
  additionalTotal: number,
  markupTotal: number,
  startRow: number,
  startCol: number,
): number {
  let row = startRow;

  // Header
  const hr = sheet.getRow(row);
  hr.getCell(startCol).value = 'PROJECT TOTAL';
  setFill(hr.getCell(startCol), DARK_BLUE);
  setFont(hr.getCell(startCol), { bold: true, size: 11, color: WHITE });
  hr.getCell(startCol).border = { top: mediumBorder(), left: mediumBorder(), bottom: thinBorder() };
  setFill(hr.getCell(startCol + 1), DARK_BLUE);
  hr.getCell(startCol + 1).border = { top: mediumBorder(), bottom: thinBorder() };
  setFill(hr.getCell(startCol + 2), DARK_BLUE);
  hr.getCell(startCol + 2).border = { top: mediumBorder(), right: mediumBorder(), bottom: thinBorder() };
  row++;

  // Discounted total
  const dr = sheet.getRow(row);
  dr.getCell(startCol).value = 'Discounted Total:';
  setFont(dr.getCell(startCol), { size: 10 });
  setFill(dr.getCell(startCol), LIGHT_GRAY);
  dr.getCell(startCol).border = { left: mediumBorder() };
  setFill(dr.getCell(startCol + 1), LIGHT_GRAY);
  dr.getCell(startCol + 2).value = discountedTotal;
  setCurrency(dr.getCell(startCol + 2));
  dr.getCell(startCol + 2).alignment = { horizontal: 'right' };
  setFill(dr.getCell(startCol + 2), LIGHT_GRAY);
  dr.getCell(startCol + 2).border = { right: mediumBorder() };
  row++;

  if (additionalTotal > 0) {
    const ar = sheet.getRow(row);
    ar.getCell(startCol).value = '+ Additional:';
    setFont(ar.getCell(startCol), { size: 10 });
    ar.getCell(startCol).border = { left: mediumBorder() };
    ar.getCell(startCol + 2).value = additionalTotal;
    setCurrency(ar.getCell(startCol + 2));
    ar.getCell(startCol + 2).alignment = { horizontal: 'right' };
    ar.getCell(startCol + 2).border = { right: mediumBorder() };
    row++;
  }

  if (markupTotal > 0) {
    const mr = sheet.getRow(row);
    mr.getCell(startCol).value = '+ Markups:';
    setFont(mr.getCell(startCol), { size: 10 });
    mr.getCell(startCol).border = { left: mediumBorder() };
    mr.getCell(startCol + 2).value = markupTotal;
    setCurrency(mr.getCell(startCol + 2));
    mr.getCell(startCol + 2).alignment = { horizontal: 'right' };
    mr.getCell(startCol + 2).border = { right: mediumBorder() };
    row++;
  }

  // Grand Total
  const grandTotal = discountedTotal + additionalTotal + markupTotal;
  const gr = sheet.getRow(row);
  gr.getCell(startCol).value = 'GRAND TOTAL:';
  setFill(gr.getCell(startCol), VERY_DARK_BLUE);
  setFont(gr.getCell(startCol), { bold: true, size: 11, color: WHITE });
  gr.getCell(startCol).border = { left: mediumBorder(), bottom: mediumBorder() };
  setFill(gr.getCell(startCol + 1), VERY_DARK_BLUE);
  gr.getCell(startCol + 1).border = { bottom: mediumBorder() };
  gr.getCell(startCol + 2).value = grandTotal;
  setCurrency(gr.getCell(startCol + 2));
  setFill(gr.getCell(startCol + 2), VERY_DARK_BLUE);
  setFont(gr.getCell(startCol + 2), { bold: true, size: 11, color: WHITE });
  gr.getCell(startCol + 2).alignment = { horizontal: 'right' };
  gr.getCell(startCol + 2).border = { right: mediumBorder(), bottom: mediumBorder() };
  row++;

  return row;
}

function writeElevationSummaryTable(
  sheet: Worksheet,
  elevations: Record<string, ElevationData>,
  settings: ProjectSettings,
  startRow: number,
  startCol: number,
): number {
  let row = startRow;

  // Check which columns to show
  const showNames = settings.show_elevation_names ?? false;
  const showQty = settings.show_elevation_quantity ?? false;
  const showDims = settings.show_elevation_dimensions ?? false;
  const showSqft = settings.show_elevation_sqft ?? false;
  const showPerimeter = settings.show_elevation_perimeter ?? false;

  if (!showNames && !showQty && !showDims && !showSqft && !showPerimeter) {
    return row;
  }

  // Title row
  const headers: string[] = [];
  if (showNames) headers.push('Elevation Name');
  if (showQty) headers.push('Quantity (EA)');
  if (showDims) headers.push('Dimensions');
  if (showSqft) headers.push('SQFT Total (SQFT)');
  if (showPerimeter) headers.push('Perimeter FT Total (FT)');

  const titleRow = sheet.getRow(row);
  titleRow.getCell(startCol).value = 'ELEVATION SUMMARY';
  setFill(titleRow.getCell(startCol), DARK_BLUE);
  setFont(titleRow.getCell(startCol), { bold: true, size: 12, color: WHITE });
  for (let i = 0; i < headers.length; i++) {
    setFill(titleRow.getCell(startCol + i), DARK_BLUE);
    titleRow.getCell(startCol + i).border = { top: mediumBorder(), bottom: thinBorder() };
  }
  row++;

  // Header row
  const headerRow = sheet.getRow(row);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(startCol + i);
    cell.value = h;
    setBoldFont(cell, 11);
    setFill(cell, LIGHT_GRAY);
    cell.border = { bottom: thinBorder() };
  });
  row++;

  // Data rows
  let totalQty = 0;
  let totalSqft = 0;
  let totalPerimeter = 0;

  for (const [name, elev] of Object.entries(elevations)) {
    const qty = elev.total_count || 1;
    const w = elev.opening_width_inches || 0;
    const h = elev.opening_height_inches || 0;
    const sqft = calculate_total_glass(w, h, qty, elev.bays_wide || 1, elev.bays_tall || 1, elev.custom_bay_widths, elev.custom_bay_heights);
    const perimeter = ((2 * (w + h)) / 12) * qty;
    totalQty += qty;
    totalSqft += sqft;
    totalPerimeter += perimeter;

    const dr = sheet.getRow(row);
    let col = startCol;
    if (showNames) dr.getCell(col++).value = name;
    if (showQty) dr.getCell(col++).value = qty;
    if (showDims) dr.getCell(col++).value = `${w.toFixed(1)}" x ${h.toFixed(1)}"`;
    if (showSqft) { dr.getCell(col).value = sqft; dr.getCell(col).numFmt = '#,##0.00'; col++; }
    if (showPerimeter) { dr.getCell(col).value = perimeter; dr.getCell(col).numFmt = '#,##0.00'; col++; }
    row++;
  }

  // Totals
  const totRow = sheet.getRow(row);
  let col = startCol;
  if (showNames) { totRow.getCell(col).value = 'TOTAL'; setBoldFont(totRow.getCell(col), 10); col++; }
  if (showQty) { totRow.getCell(col).value = totalQty; setBoldFont(totRow.getCell(col), 10); col++; }
  if (showDims) { totRow.getCell(col).value = ''; col++; }
  if (showSqft) { totRow.getCell(col).value = totalSqft; totRow.getCell(col).numFmt = '#,##0.00'; setBoldFont(totRow.getCell(col), 10); col++; }
  if (showPerimeter) { totRow.getCell(col).value = totalPerimeter; totRow.getCell(col).numFmt = '#,##0.00'; setBoldFont(totRow.getCell(col), 10); col++; }
  for (let i = startCol; i < col; i++) {
    setFill(totRow.getCell(i), LIGHT_GRAY);
    totRow.getCell(i).border = { top: thinBorder() };
  }
  row += 2;

  return row;
}

// ---------------------------------------------------------------------------
// Bay Diagram Generation (Canvas-based, matching Python _create_bay_diagram)
// ---------------------------------------------------------------------------

/**
 * Parse a door size string like "3' X 7'" into [widthInches, heightInches].
 */
function parseDoorSizeInches(sizeStr: string): [number, number] {
  const m = sizeStr.match(/(\d+)'\s*[xX]\s*(\d+)'/);
  if (m) return [parseFloat(m[1]) * 12, parseFloat(m[2]) * 12];
  return [36, 84]; // default 3'x7'
}

/**
 * Creates a bay distribution diagram as a PNG base64 string for embedding in Excel.
 * @param mode 'cl' for centerline dimensions, 'dlo' for D.L.O. dimensions
 */
function createBayDiagram(
  baysWide: number,
  baysTall: number,
  openingWidth: number,
  openingHeight: number,
  customBayWidths?: number[],
  customBayHeights?: number[],
  doors?: DoorConfig[],
  mode: 'cl' | 'dlo' = 'cl',
): string | null {
  if (baysWide <= 0 || baysTall <= 0 || openingWidth <= 0 || openingHeight <= 0) return null;

  const bayWidths =
    customBayWidths && customBayWidths.length === baysWide
      ? customBayWidths
      : Array(baysWide).fill(openingWidth / baysWide) as number[];

  const bayHeights =
    customBayHeights && customBayHeights.length === baysTall
      ? customBayHeights
      : Array(baysTall).fill(openingHeight / baysTall) as number[];

  const diagramWidth = 400;
  const diagramHeight = 300;
  const margin = doors && doors.length > 0 ? 15 : 20;

  const canvas = document.createElement('canvas');
  canvas.width = diagramWidth;
  canvas.height = diagramHeight;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;

  // White background
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, diagramWidth, diagramHeight);

  const maxDisplayWidth = diagramWidth - 2 * margin;
  const maxDisplayHeight = diagramHeight - 2 * margin - 60;

  const totalWidth = bayWidths.reduce((s, w) => s + w, 0);
  const totalHeight = bayHeights.reduce((s, h) => s + h, 0);
  const scaleX = totalWidth > 0 ? maxDisplayWidth / totalWidth : 1;
  const scaleY = totalHeight > 0 ? maxDisplayHeight / totalHeight : 1;
  const scale = Math.min(scaleX, scaleY);

  const scaledTotalWidth = totalWidth * scale;
  const scaledTotalHeight = totalHeight * scale;
  const startX = margin + (maxDisplayWidth - scaledTotalWidth) / 2;
  const startY = margin + 30;

  // Title
  ctx.fillStyle = '#000000';
  ctx.font = 'bold 12px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(
    mode === 'dlo' ? 'Bay Distribution — D.L.O. Dimensions' : 'Bay Distribution — C/L Dimensions',
    diagramWidth / 2,
    10,
  );

  // Draw vertical grid lines (between bays)
  let accX = startX;
  for (let i = 0; i < baysWide; i++) {
    if (i > 0) {
      ctx.beginPath();
      ctx.moveTo(accX, startY);
      ctx.lineTo(accX, startY + scaledTotalHeight);
      ctx.strokeStyle = '#808080';
      ctx.lineWidth = 2;
      ctx.stroke();
    }
    accX += bayWidths[i] * scale;
  }

  // Draw horizontal grid lines (between bays)
  let accY = startY;
  for (let i = 0; i < baysTall; i++) {
    if (i > 0) {
      ctx.beginPath();
      ctx.moveTo(startX, accY);
      ctx.lineTo(startX + scaledTotalWidth, accY);
      ctx.strokeStyle = '#808080';
      ctx.lineWidth = 2;
      ctx.stroke();
    }
    accY += bayHeights[i] * scale;
  }

  // Outer rectangle
  ctx.strokeStyle = '#000000';
  ctx.lineWidth = 3;
  ctx.strokeRect(startX, startY, scaledTotalWidth, scaledTotalHeight);

  // Bay labels (B1, B2, ... with dimensions — C/L or D.L.O.)
  ctx.font = '8px Arial, sans-serif';
  ctx.fillStyle = '#000000';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  let bayNum = 1;
  let rowY = startY;
  for (let row = 0; row < baysTall; row++) {
    let colX = startX;
    // SVG draws top-to-bottom: row 0 at top = data row (baysTall-1) which is the top row
    // In our data, row 0 = bottom. So dataRow for SVG row index:
    const dataRow = baysTall - 1 - row;
    for (let col = 0; col < baysWide; col++) {
      const cx = colX + (bayWidths[col] * scale) / 2;
      const cy = rowY + (bayHeights[row] * scale) / 2;
      ctx.fillText(`B${bayNum}`, cx, cy - 6);
      if (mode === 'dlo') {
        const dloW = calculateDloWidth(bayWidths[col], col, baysWide);
        const dloH = calculateDloHeight(bayHeights[row], dataRow, baysTall);
        ctx.fillStyle = '#1565C0';
        ctx.fillText(`${dloW.toFixed(1)}" x ${dloH.toFixed(1)}"`, cx, cy + 6);
        ctx.fillStyle = '#000000';
      } else {
        ctx.fillText(`${bayWidths[col].toFixed(1)}" x ${bayHeights[row].toFixed(1)}"`, cx, cy + 6);
      }
      colX += bayWidths[col] * scale;
      bayNum++;
    }
    rowY += bayHeights[row] * scale;
  }

  // Draw door bands (green, bottom-aligned)
  const widthRef = totalWidth > 0 ? totalWidth : openingWidth;
  if (doors && doors.length > 0 && widthRef > 0 && openingHeight > 0) {
    for (const door of doors) {
      const [doorW, doorH] = parseDoorSizeInches(door.size);
      const count = door.count || 1;
      for (let c = 0; c < count; c++) {
        let xCenter: number | undefined;
        if (door.x_positions && door.x_positions[c] != null) {
          xCenter = door.x_positions[c];
        } else if (door.x_in != null && count === 1) {
          xCenter = door.x_in;
        }
        if (xCenter == null) continue;

        const leftIn = Math.max(0, Math.min(xCenter - doorW / 2, openingWidth));
        const rightIn = Math.max(0, Math.min(xCenter + doorW / 2, openingWidth));
        if (rightIn <= leftIn) continue;

        const pxLeft = startX + (leftIn / widthRef) * scaledTotalWidth;
        const pxRight = startX + (rightIn / widthRef) * scaledTotalWidth;
        const doorHPx = (doorH / openingHeight) * scaledTotalHeight;
        const pxBottom = startY + scaledTotalHeight;
        const pxTop = pxBottom - doorHPx;

        ctx.fillStyle = '#A5D6A7';
        ctx.fillRect(pxLeft, pxTop, pxRight - pxLeft, pxBottom - pxTop);
        ctx.strokeStyle = '#2E7D32';
        ctx.lineWidth = 2;
        ctx.strokeRect(pxLeft, pxTop, pxRight - pxLeft, pxBottom - pxTop);
      }
    }
  }

  // Total dimensions at bottom
  ctx.fillStyle = '#000000';
  ctx.font = '8px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(
    `Total (C/L): ${openingWidth.toFixed(1)}" W x ${openingHeight.toFixed(1)}" H`,
    diagramWidth / 2,
    diagramHeight - 20,
  );

  const dataUrl = canvas.toDataURL('image/png');
  return dataUrl.replace(/^data:image\/png;base64,/, '');
}

// ---------------------------------------------------------------------------
// Pie Chart Generation (Canvas-based, matching Python PIL implementation)
// ---------------------------------------------------------------------------

interface PieSegment {
  name: string;
  value: number;
  pct: number;
  color: string;
}

/**
 * Creates a cost breakdown pie chart as a PNG base64 string.
 * Segments: Active Materials (blue), Additional (green), Profit/Markups (purple), Residual/Waste (orange)
 * Matches Python's _create_cost_pie_chart() exactly.
 */
function createCostPieChart(
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

  // Create canvas (works in browser context)
  const canvas = document.createElement('canvas');
  canvas.width = chartWidth;
  canvas.height = chartHeight;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;

  // White background
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, chartWidth, chartHeight);

  // Calculate percentages
  const materialPct = (materialCost / grandTotal) * 100;
  const miscPct = (miscCost / grandTotal) * 100;
  const markupPct = (markupCost / grandTotal) * 100;
  const residualPct = (residualCost / grandTotal) * 100;

  // Segment colors (matching Python)
  const MATERIAL_COLOR = '#4472C4';  // Blue
  const MISC_COLOR = '#548235';      // Green
  const MARKUP_COLOR = '#7030A0';    // Purple
  const RESIDUAL_COLOR = '#ED7D31';  // Orange

  // Build segments (only non-zero values drawn as slices)
  const segments: PieSegment[] = [];
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

  // Draw pie slices
  let startAngle = -Math.PI / 2; // Start at 12 o'clock
  for (const seg of segments) {
    if (seg.pct <= 0) continue;
    const sweepAngle = (seg.pct / 100) * 2 * Math.PI;
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, startAngle, startAngle + sweepAngle);
    ctx.closePath();
    ctx.fillStyle = seg.color;
    ctx.fill();
    // White border between slices
    ctx.strokeStyle = '#FFFFFF';
    ctx.lineWidth = 2;
    ctx.stroke();
    startAngle += sweepAngle;
  }

  // Legend at bottom
  const legendY = 270;
  const legendBoxSize = 12;
  const legendSpacing = 22;

  const legendItems: PieSegment[] = [
    { name: 'Active Materials', value: materialCost, pct: materialPct, color: MATERIAL_COLOR },
    { name: 'Additional', value: miscCost, pct: miscPct, color: MISC_COLOR },
    { name: 'Profit/Markups', value: markupCost, pct: markupPct, color: MARKUP_COLOR },
    { name: 'Residual/Waste', value: residualCost, pct: residualPct, color: RESIDUAL_COLOR },
  ];

  for (let i = 0; i < legendItems.length; i++) {
    const item = legendItems[i];
    const yPos = legendY + (i * legendSpacing);

    // Color box
    ctx.fillStyle = item.color;
    ctx.fillRect(30, yPos, legendBoxSize, legendBoxSize);
    ctx.strokeStyle = '#333333';
    ctx.lineWidth = 1;
    ctx.strokeRect(30, yPos, legendBoxSize, legendBoxSize);

    // Label
    ctx.fillStyle = '#333333';
    ctx.font = '10px Arial, sans-serif';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    const label = `${item.name}: $${item.value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${item.pct.toFixed(1)}%)`;
    ctx.fillText(label, 50 + legendBoxSize, yPos + legendBoxSize / 2);
  }

  // Grand total at bottom
  ctx.fillStyle = '#333333';
  ctx.font = '9px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(
    `Grand Total: $${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    centerX,
    chartHeight - 15,
  );

  // Convert to base64 PNG (strip the data URL prefix for ExcelJS)
  const dataUrl = canvas.toDataURL('image/png');
  return dataUrl.replace(/^data:image\/png;base64,/, '');
}

// ---------------------------------------------------------------------------
// Main export function
// ---------------------------------------------------------------------------

export async function exportToExcel(
  projectName: string,
  elevations: Record<string, ElevationData>,
  doors: Record<string, DoorConfig[]>,
  settings: ProjectSettings,
  materials: Record<string, ExtraMaterial>,
  reportConfig?: ReportConfig,
): Promise<void> {
  const ExcelJS = (await import('exceljs')).default;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'United Glass Ventures - Estimation Tool';
  workbook.created = new Date();

  // ========================================================================
  // Pre-pass: compute running grand total for discount multiplier tier
  // Re-price every item from scratch using getPriceByPart (summary=true)
  // so the total is accurate regardless of stale stored prices.
  // This matches the Python create_summary_sheet Step 1 approach.
  // ========================================================================

  let runningGrandTotal = 0;
  for (const [, elev] of Object.entries(elevations)) {
    if (!elev.calculated_outputs) continue;
    const elevFinish = elev.finish || '';
    for (const output of elev.calculated_outputs) {
      const cat = classifyOutput(output);
      if (cat === 'calculations') continue;
      if (output.manual || cat === 'glass' || cat === 'fabrication' || cat === 'doors') {
        // Manual items: output.price is already the total cost
        runningGrandTotal += output.price ?? 0;
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
  const threshold = settings.discount_threshold ?? 50000;

  // ========================================================================
  // Per-elevation sheets
  // ========================================================================

  const sortedElevNames = Object.keys(elevations).sort();

  for (const elevName of sortedElevNames) {
    const elev = elevations[elevName];
    if (!elev.calculated_outputs || elev.calculated_outputs.length === 0) continue;

    const sheetName = elevName.replace(/[\\/*?:\[\]]/g, '_').slice(0, 31);
    const sheet = workbook.addWorksheet(sheetName);

    const totalCount = elev.total_count || 1;
    const showPerElev = totalCount > 1;
    const elevDoors = doors[elevName] || [];

    // Column widths — initial defaults (autofit runs after data is written)
    sheet.getColumn(1).width = 22;
    sheet.getColumn(2).width = 22;
    sheet.getColumn(3).width = 4; // spacer
    sheet.getColumn(4).width = 4; // spacer
    // Material section columns start at 5
    for (let c = 5; c <= 16; c++) {
      sheet.getColumn(c).width = 18;
    }

    // Per-elevation section config
    const elevSections = reportConfig?.per_elevation_sections?.[elevName];

    // --- System Input (cols A-B, rows 1-15) ---
    let inputEndRow = 1;
    if (elevSections?.system_input !== false) {
      inputEndRow = writeSystemInput(sheet, elevName, elev, elevDoors, 1);
    }

    // --- Material sections (starting col E) ---
    const categories = buildElevationCategories(elev.calculated_outputs, elev.finish, totalCount, multiplier, elev.single_elevation_outputs);
    const startCol = 5;
    let sectionRow = 1;

    const catOrder: [string, string][] = [
      ['profiles', 'Profiles'],
      ['accessories', 'Accessories'],
      ['gaskets', 'Gaskets'],
      ['doors', 'Doors'],
      ['glass', 'Glass'],
      ['fabrication', 'Labor'],
    ];

    for (const [key, label] of catOrder) {
      // Skip section if unchecked in stock list
      if (elevSections?.[key] === false) continue;
      sectionRow = writeMaterialSection(sheet, label, categories[key], totalCount, showPerElev, sectionRow, startCol);
    }

    // --- Elevation Cost Summary ---
    if (elevSections?.elevation_cost_summary !== false) {
      sectionRow = writeElevationCostSummary(sheet, categories, elevName, totalCount, sectionRow, startCol, elevSections);
    }

    // Auto-fit all material columns so descriptions are never truncated
    autofitColumns(sheet, 5, 16);

    // --- Bay Diagrams (C/L above, D.L.O. below) ---
    if (elevSections?.diagram !== false && elev.bays_wide > 0 && elev.bays_tall > 0) {
      const noteRow = inputEndRow;
      const noteCell = sheet.getRow(noteRow).getCell(1);
      noteCell.value = '*Bay Distribution — C/L & D.L.O. Diagrams';
      setFont(noteCell, { size: 12 });

      const diagramRow = noteRow + 1;
      // 300px diagram height ÷ ~20px per default row ≈ 15 rows, +1 for spacing
      const diagramRowSpan = 16;

      // C/L diagram (top)
      const clBase64 = createBayDiagram(
        elev.bays_wide, elev.bays_tall,
        elev.opening_width_inches, elev.opening_height_inches,
        elev.custom_bay_widths, elev.custom_bay_heights,
        elevDoors, 'cl',
      );
      if (clBase64) {
        const clImg = workbook.addImage({ base64: clBase64, extension: 'png' });
        sheet.addImage(clImg, {
          tl: { col: 0, row: diagramRow - 1 },
          ext: { width: 400, height: 300 },
        });
      }

      // D.L.O. diagram (directly below C/L)
      const dloBase64 = createBayDiagram(
        elev.bays_wide, elev.bays_tall,
        elev.opening_width_inches, elev.opening_height_inches,
        elev.custom_bay_widths, elev.custom_bay_heights,
        elevDoors, 'dlo',
      );
      if (dloBase64) {
        const dloImg = workbook.addImage({ base64: dloBase64, extension: 'png' });
        sheet.addImage(dloImg, {
          tl: { col: 0, row: diagramRow - 1 + diagramRowSpan },
          ext: { width: 400, height: 300 },
        });
      }
    }
  }

  // ========================================================================
  // Summary sheet
  // ========================================================================

  const summaryIncluded = reportConfig?.summary_included !== false;

  if (summaryIncluded) {
  const summarySheet = workbook.addWorksheet('Summary');
  // 10 data columns + extra for cost overview side placement
  for (let c = 1; c <= 14; c++) {
    summarySheet.getColumn(c).width = c <= 2 ? 28 : c <= 5 ? 22 : 18;
  }

  const sumSections = reportConfig?.summary_options?.sections;
  const sumCostOverview = reportConfig?.summary_options?.cost_overview;

  // Build summary categories
  const summaryCategories = buildSummaryCategories(elevations, materials, multiplier);

  // Compute totals
  let totalListPrice = 0;
  let totalDiscountedPrice = 0;
  let totalResidual = 0;
  const catKeys = ['profiles', 'accessories', 'gaskets', 'doors', 'glass', 'fabrication'];
  for (const key of catKeys) {
    totalListPrice += summaryCategories[key].total_original;
    totalDiscountedPrice += summaryCategories[key].total_discounted;
    totalResidual += summaryCategories[key].total_residual;
  }

  const wastePct = totalDiscountedPrice > 0 ? (totalResidual / totalDiscountedPrice) * 100 : 0;

  // Write category sections (all 10 columns per category, matching Python)
  let currentRow = 1;
  const summaryCatOrder: [string, string][] = [
    ['profiles', 'Profiles'],
    ['accessories', 'Accessories'],
    ['gaskets', 'Gaskets'],
    ['doors', 'Doors'],
    ['glass', 'Glass'],
    ['fabrication', 'Labor'],
  ];

  for (const [key, label] of summaryCatOrder) {
    // Skip section if unchecked in stock list
    if (sumSections?.[key] === false) continue;
    currentRow = writeSummaryCategorySection(summarySheet, label, key, summaryCategories[key], currentRow, 1);
  }

  // Elevation Summary Table
  currentRow = writeElevationSummaryTable(summarySheet, elevations, settings, currentRow, 1);

  // Cost Overview Box — track start row for diagram placement
  const costOverviewStartRow = currentRow + 1;
  currentRow += 1;
  currentRow = writeCostOverviewBox(summarySheet, totalListPrice, totalDiscountedPrice, totalResidual, wastePct, currentRow, 1);

  // Additional Costs
  let additionalTotal = 0;
  if (sumCostOverview?.additional_costs !== false) {
    currentRow += 1;
    const result = writeAdditionalCostsSection(
      summarySheet, settings, totalDiscountedPrice, currentRow, 1,
    );
    currentRow = result.row;
    additionalTotal = result.total;
  }

  // Markups
  let markupTotal = 0;
  if (sumCostOverview?.markups !== false) {
    currentRow += 1;
    const result = writeMarkupsSection(
      summarySheet, settings, summaryCategories, totalDiscountedPrice, totalResidual, currentRow, 1,
    );
    currentRow = result.row;
    markupTotal = result.total;
  }

  // Project Total
  if (sumCostOverview?.project_total !== false) {
    currentRow += 1;
    currentRow = writeProjectTotalSection(summarySheet, totalDiscountedPrice, additionalTotal, markupTotal, currentRow, 1);
  }

  // Auto-fit summary columns so descriptions are never truncated
  autofitColumns(summarySheet, 1, 14);

  // ========================================================================
  // Pie Chart - Cost Distribution (adjacent to Cost Overview, right side)
  // ========================================================================

  if (sumCostOverview?.diagram !== false) {
  try {
    const activeMaterialCost = Math.max(0, totalDiscountedPrice - totalResidual);
    const pieChartBase64 = createCostPieChart(
      activeMaterialCost,
      additionalTotal,
      markupTotal,
      totalResidual,
    );

    if (pieChartBase64) {
      const imageId = workbook.addImage({
        base64: pieChartBase64,
        extension: 'png',
      });

      // Place adjacent to cost overview (column D/E, same row as cost overview start)
      summarySheet.addImage(imageId, {
        tl: { col: 3, row: costOverviewStartRow - 1 },  // 0-based: col D = 3
        ext: { width: 380, height: 360 },
      });
    }
  } catch (e) {
    console.warn('Could not generate pie chart:', e);
  }
  } // end diagram check
  } // end summaryIncluded

  // ========================================================================
  // Write buffer and trigger download
  // ========================================================================

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);

  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `${projectName.replace(/[^a-zA-Z0-9_-]/g, '_')}_Estimate.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();

  setTimeout(() => {
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
  }, 100);
}
