import { PART_NUMBER_MAP } from '@/data/part-number';
import * as formulas from '@/lib/formulas';
import { CalculatedOutput, DoorConfig } from '@/types';

/**
 * Calculates all output quantities for the 'YES 45TU Center Set' system.
 * Returns a list of objects with description, quantity, part number, and type.
 *
 * Key differences from Front Set (OG):
 * - BE9-2553 is used for ALL verticals (jambs + intermediates) AND head pieces
 * - Head = baysWide pieces at bay width (not one full-width piece)
 * - No shear blocks (E1-1058/1059) or their screws (PC-1028/FC-1212/PC-1210)
 * - No PC-1216 short spline screw
 * - No E2-0611 inside setting block
 * - DLO uses uniform 8/3" deduction per bay (no edge/interior distinction)
 */

// ---------------------------------------------------------------------------
// Center-set DLO constants
// In center set every vertical contributes the same clearance to each adjacent
// bay, yielding a uniform 8/3" deduction (no edge/interior distinction).
// ---------------------------------------------------------------------------

const CS_DLO_DEDUCTION = 8 / 3;       // inches — uniform per bay
const CS_SILL_DEDUCTION = 7 / 16;     // inches — extra at bottom row
const CS_GLASS_MAKE_ADDITION = 3 / 4; // inches — glass make = DLO + 3/4"

function csDloWidth(bayWidth: number): number {
  return bayWidth - CS_DLO_DEDUCTION;
}

function csDloHeight(bayHeight: number, isBottom: boolean): number {
  let dlo = bayHeight - CS_DLO_DEDUCTION;
  if (isBottom) dlo -= CS_SILL_DEDUCTION;
  return dlo;
}

// ---------------------------------------------------------------------------
// Center-set glass & gasket calculations
// ---------------------------------------------------------------------------

function calculateCsGasket(
  openingWidth: number,
  openingHeight: number,
  totalCount: number,
  baysWide: number,
  baysTall: number,
  customBayWidths?: number[],
  customBayHeights?: number[],
): number {
  const bayWidths =
    customBayWidths && customBayWidths.length === baysWide
      ? customBayWidths
      : Array(baysWide).fill(openingWidth / baysWide);
  const bayHeights =
    customBayHeights && customBayHeights.length === baysTall
      ? customBayHeights
      : Array(baysTall).fill(openingHeight / baysTall);

  let totalInches = 0;
  for (let col = 0; col < baysWide; col++) {
    const glassW = csDloWidth(bayWidths[col]) + CS_GLASS_MAKE_ADDITION;
    for (let row = 0; row < baysTall; row++) {
      const glassH = csDloHeight(bayHeights[row], row === 0) + CS_GLASS_MAKE_ADDITION;
      // 2 sides × perimeter of this lite
      totalInches += 2 * 2 * (glassW + glassH);
    }
  }
  return (totalInches * totalCount) / 12;
}

// ---------------------------------------------------------------------------
// Center-set profile helpers
// ---------------------------------------------------------------------------

function calculateCsVerticalBe92553(
  openingHeight: number,
  totalCount: number,
  baysWide: number,
): number[] {
  // 2 jambs + (baysWide - 1) intermediates = baysWide + 1 pieces
  // Cut to height - 7/16" (verticals sit on top of sill flashing)
  const pieceFt = (openingHeight - CS_SILL_DEDUCTION) / 12;
  return Array((baysWide + 1) * totalCount).fill(pieceFt);
}

function calculateCsHeadBe92553(
  openingWidth: number,
  totalCount: number,
  baysWide: number,
  customBayWidths?: number[],
): number[] {
  // baysWide pieces, each cut to DLO width (bay_width - 8/3")
  // Horizontals fit between vertical profiles
  const bayWidthsFt =
    customBayWidths && customBayWidths.length === baysWide
      ? customBayWidths.map(w => (w - CS_DLO_DEDUCTION) / 12)
      : Array(baysWide).fill((openingWidth / baysWide - CS_DLO_DEDUCTION) / 12);
  const result: number[] = [];
  for (let i = 0; i < totalCount; i++) result.push(...bayWidthsFt);
  return result;
}

function calculateCsIntHorizontal(
  openingWidth: number,
  totalCount: number,
  baysWide: number,
  baysTall: number,
  customBayWidths?: number[],
): number[] {
  // Each piece cut to DLO width (bay_width - 8/3") — fits between verticals
  const bayWidthsFt =
    customBayWidths && customBayWidths.length === baysWide
      ? customBayWidths.map(w => (w - CS_DLO_DEDUCTION) / 12)
      : Array(baysWide).fill((openingWidth / baysWide - CS_DLO_DEDUCTION) / 12);
  const rows = baysTall - 1;
  const pieces: number[] = [];
  for (let i = 0; i < rows; i++) pieces.push(...bayWidthsFt);
  const result: number[] = [];
  for (let i = 0; i < totalCount; i++) result.push(...pieces);
  return result;
}

// ---------------------------------------------------------------------------
// Center-set sill / flashing / glass-stop overrides
// Override shared formulas because center-set cut lengths differ:
//   • Verticals/fillers: height − 7/16" (sill flashing occupies that space)
//   • Sill flashing: width + 1/4" (extends past jamb profiles)
//   • Sill / glass-stop horizontals: DLO width = bay_width − 8/3"
// ---------------------------------------------------------------------------

function calculateCsFlushFillerV(
  baysWide: number,
  totalCount: number,
  openingHeight: number,
): number[] {
  // (baysWide - 1) pieces, each cut to height - 7/16"
  const pieceFt = (openingHeight - CS_SILL_DEDUCTION) / 12;
  return Array((baysWide - 1) * totalCount).fill(pieceFt);
}

function calculateCsSillFtH(
  openingWidth: number,
  totalCount: number,
  baysWide: number,
  customBayWidths?: number[],
): number[] {
  // baysWide pieces per elevation, each cut to DLO width (bay_width - 8/3")
  const bayWidthsFt =
    customBayWidths && customBayWidths.length === baysWide
      ? customBayWidths.map(w => (w - CS_DLO_DEDUCTION) / 12)
      : Array(baysWide).fill((openingWidth / baysWide - CS_DLO_DEDUCTION) / 12);
  const result: number[] = [];
  for (let i = 0; i < totalCount; i++) result.push(...bayWidthsFt);
  return result;
}

function calculateCsSillFlashingH(
  openingWidth: number,
  totalCount: number,
): number[] {
  // 1 piece per elevation at opening_width + 1/4" (extends past jambs)
  const pieceFt = (openingWidth + 0.25) / 12;
  return Array(totalCount).fill(pieceFt);
}

function calculateCsGlassStop(
  openingWidth: number,
  baysTall: number,
  totalCount: number,
  baysWide: number,
  customBayWidths?: number[],
): number[] {
  // baysWide * baysTall pieces, each cut to DLO width (bay_width - 8/3")
  const bayWidthsFt =
    customBayWidths && customBayWidths.length === baysWide
      ? customBayWidths.map(w => (w - CS_DLO_DEDUCTION) / 12)
      : Array(baysWide).fill((openingWidth / baysWide - CS_DLO_DEDUCTION) / 12);
  const piecesPerElev: number[] = [];
  for (let i = 0; i < baysTall; i++) piecesPerElev.push(...bayWidthsFt);
  const result: number[] = [];
  for (let i = 0; i < totalCount; i++) result.push(...piecesPerElev);
  return result;
}

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------

export function calculateYes45tuCenterSetQuantities(
  baysWide: number,
  baysTall: number,
  totalCount: number,
  openingWidth: number,
  openingHeight: number,
  doors?: DoorConfig[],
  customBayWidths?: number[],
  customBayHeights?: number[],
  glassPerSqft?: number,
  fabricationCostPerJoint?: number
): CalculatedOutput[] {
  if (!doors) doors = [];

  // --- Accessory quantities ---
  const waterDeflector = 2 * (baysTall - 1) * baysWide * totalCount;
  // PC-1220: all horizontal members (head + int + sill) = baysWide*(baysTall+1), 4 screws each
  const screwPC1220 = 4 * baysWide * (baysTall + 1) * totalCount;
  // E2-0020: 2 per lite
  const settingBlock = 2 * baysWide * baysTall * totalCount;
  // E2-0153: 2 per lite
  const antiWalkBlock = 2 * baysWide * baysTall * totalCount;
  // E1-1054: sill(3*baysWide) + head(2*(baysWide+1)) + jambs(2*(baysTall+1))
  const flatFiller = (5 * baysWide + 2 * baysTall + 4) * totalCount;

  const outputs: [string, number | number[]][] = [
    // --- Accessories ---
    ['E1-0199', formulas.calculateEndDam(totalCount)],
    ['E2-0047', waterDeflector],
    ['PC-1220', screwPC1220],
    ['PM-1008-SS', formulas.calculateSillFlashScrew(baysWide, totalCount)],
    ['UA-1212', formulas.calculateEndDamScrew(totalCount)],
    ['E2-0020', settingBlock],
    ['E2-0153', antiWalkBlock],
    ['E1-1054', flatFiller],
    // --- Profiles ---
    // BE9-2553 vertical: jambs + intermediate verticals (baysWide + 1 pieces)
    ['BE9-2553', calculateCsVerticalBe92553(openingHeight, totalCount, baysWide)],
    // BE9-2552: shallow pocket filler verticals (baysWide - 1 pieces)
    ['BE9-2552', calculateCsFlushFillerV(baysWide, totalCount, openingHeight)],
    // BE9-2553 head: baysWide pieces at bay width
    ['BE9-2553', calculateCsHeadBe92553(openingWidth, totalCount, baysWide, customBayWidths)],
    // BE9-2579: sill pieces at DLO width
    ['BE9-2579', calculateCsSillFtH(openingWidth, totalCount, baysWide, customBayWidths)],
    // BE9-2556: intermediate horizontals
    ['BE9-2556', calculateCsIntHorizontal(openingWidth, totalCount, baysWide, baysTall, customBayWidths)],
    // BE9-2578: sill flashing at opening_width + 1/4"
    ['BE9-2578', calculateCsSillFlashingH(openingWidth, totalCount)],
    // E9-1015: glass stop at DLO width per lite
    ['E9-1015', calculateCsGlassStop(openingWidth, baysTall, totalCount, baysWide, customBayWidths)],
    // E2-0052: glazing gasket
    ['E2-0052', calculateCsGasket(openingWidth, openingHeight, totalCount, baysWide, baysTall, customBayWidths, customBayHeights)],
  ];

  // --- Build bay dimension arrays for per-pane glass ---
  const bayWidths: number[] =
    customBayWidths && customBayWidths.length === baysWide
      ? customBayWidths
      : Array(baysWide).fill(openingWidth / baysWide);

  const bayHeights: number[] =
    customBayHeights && customBayHeights.length === baysTall
      ? customBayHeights
      : Array(baysTall).fill(openingHeight / baysTall);

  // Group panes by unique center-set DLO dimensions
  const paneGroups = new Map<string, { width: number; height: number; count: number }>();
  for (let col = 0; col < baysWide; col++) {
    for (let row = 0; row < baysTall; row++) {
      const w = csDloWidth(bayWidths[col]);
      const h = csDloHeight(bayHeights[row], row === 0);
      const key = `${w.toFixed(4)}_${h.toFixed(4)}`;
      const existing = paneGroups.get(key);
      if (existing) {
        existing.count += 1;
      } else {
        paneGroups.set(key, { width: w, height: h, count: 1 });
      }
    }
  }

  // Door area calculations
  const hasDoors = doors && doors.length > 0;
  let totalDoorArea = 0;
  let totalGlassToAddBack = 0;

  if (hasDoors) {
    const doorsWithTotalCount: DoorConfig[] = doors.map(door => ({
      ...door,
      count: (door.count ?? 0) * totalCount,
    }));
    totalDoorArea = formulas.calculateTotalDoorArea(doorsWithTotalCount);
    totalGlassToAddBack = formulas.calculateGlassToAddBack(doorsWithTotalCount);
  }

  const results: CalculatedOutput[] = [];

  // BE9-2553 appears twice (vertical + head); track counter for labeling
  let be9_2553Counter = 0;

  for (const [partNumber, quantity] of outputs) {
    let desc: string | null = null;
    let partType: string | null = null;
    let resolvedPartNumber = partNumber;

    if (partNumber === 'BE9-2553') {
      be9_2553Counter++;
      const baseDesc = PART_NUMBER_MAP['profiles']?.['BE9-2553'] ?? 'UNKNOWN';
      desc = be9_2553Counter === 1
        ? `Vertical ${baseDesc}`
        : `Horizontal ${baseDesc} (Head)`;
      partType = 'profiles';
    } else {
      for (const [category, partsDict] of Object.entries(PART_NUMBER_MAP)) {
        if (partNumber in partsDict) {
          const baseDesc = partsDict[partNumber];
          if (['BE9-2556', 'E9-1015'].includes(partNumber)) {
            desc = `Horizontal ${baseDesc}`;
          } else if (partNumber === 'BE9-2552') {
            desc = `Vertical ${baseDesc}`;
          } else if (partNumber === 'BE9-2579') {
            desc = `Horizontal ${baseDesc} (Sill)`;
          } else {
            desc = baseDesc;
          }
          partType = category;
          break;
        }
      }
    }

    if (desc === null) {
      desc = 'UNKNOWN';
      partType = 'UNKNOWN';
      resolvedPartNumber = 'UNKNOWN';
    }

    results.push({
      description: desc,
      quantity,
      part_number: resolvedPartNumber,
      type: partType!,
    });
  }

  // --- Per-pane glass outputs (center-set uses glass make = DLO + 3/4") ---
  const glassRate = glassPerSqft != null ? Number(glassPerSqft) : 10.5;
  const doorDeduction = totalDoorArea - totalGlassToAddBack;
  const totalPaneArea = Array.from(paneGroups.values()).reduce(
    (sum, g) => {
      const glassW = g.width + CS_GLASS_MAKE_ADDITION;
      const glassH = g.height + CS_GLASS_MAKE_ADDITION;
      return sum + (glassW * glassH / 144) * g.count * totalCount;
    }, 0,
  );
  const adjustedGlassArea = Math.max(totalPaneArea - doorDeduction, 0);

  const glassOutputs: CalculatedOutput[] = [];
  if (adjustedGlassArea === 0 && hasDoors) {
    glassOutputs.push({
      description: 'Glass Area',
      quantity: 0,
      part_number: 'N/A',
      type: 'Glass',
      price: 0,
      unit: 'sqft',
      manual: true,
      message: 'Total door area equals or exceeds total glass area. No glass is needed.',
    });
  } else {
    for (const [, group] of paneGroups) {
      const glassW = group.width + CS_GLASS_MAKE_ADDITION;
      const glassH = group.height + CS_GLASS_MAKE_ADDITION;
      const paneAreaSqft = (glassW * glassH) / 144;
      const totalPanes = group.count * totalCount;
      glassOutputs.push({
        description: `Glass Pane — DLO: ${group.width.toFixed(2)}" × ${group.height.toFixed(2)}"`,
        quantity: totalPanes,
        part_number: 'N/A',
        type: 'Glass',
        price: glassRate,
        unit: 'panes',
        area_sqft: paneAreaSqft,
        manual: true,
      });
    }

    if (hasDoors && doorDeduction > 0) {
      glassOutputs.push({
        description: 'Glass — Door Area Deduction',
        quantity: -doorDeduction,
        part_number: 'N/A',
        type: 'Glass',
        price: glassRate,
        unit: 'sqft',
        manual: true,
      });
    }
  }

  const fabPrice = fabricationCostPerJoint != null ? Number(fabricationCostPerJoint) : 15.0;
  const manualOutputs: CalculatedOutput[] = [
    ...glassOutputs,
    {
      description: 'Joints Fabrication Labor',
      quantity: formulas.calculateFabricationJoints(baysWide, baysTall, totalCount),
      part_number: 'N/A',
      type: 'Fabrication',
      price: fabPrice,
      unit: 'joints',
      manual: true,
    },
  ];

  if (hasDoors) {
    manualOutputs.push({
      description: 'Door Area (to subtract from glass)',
      quantity: totalDoorArea,
      part_number: 'N/A',
      type: 'Calculations',
      unit: 'sqft',
      manual: true,
    });
  }

  results.push(...manualOutputs);
  return results;
}
