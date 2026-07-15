import { PART_NUMBER_MAP } from '@/data/part-number';
import * as formulas from '@/lib/formulas';
import { CalculatedOutput, DoorConfig } from '@/types';

/**
 * Calculates all the specific output quantities for the 'YES 45TU Front Set(OG)' system
 * by calling dedicated formula functions.
 * Returns a list of objects with description, quantity, part number, and type.
 */
export function calculateYes45tuQuantities(
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
  // Safety check for doors
  if (!doors) {
    doors = [];
  }

  const outputs: [string, number | number[]][] = [
    ['E1-0199', formulas.calculateEndDam(totalCount)],
    ['E2-0047', formulas.calculateWaterDeflector(baysWide, totalCount)],
    ['PC-1220', formulas.calculateAssemblyScrew(baysWide, baysTall, totalCount)],
    ['PM-1006-SS', formulas.calculateSillFlashScrew(baysWide, totalCount)],
    ['UA-1212', formulas.calculateEndDamScrew(totalCount)],
    ['E1-2530', formulas.calculateSettingBlockChair(baysWide, totalCount)],
    ['E2-0166', formulas.calculateSideBlock(baysWide, baysTall, totalCount)],
    ['E2-0177', formulas.calculateSettingBlock(baysWide, totalCount)],
    ['E2-0545', formulas.calculateAntiWalkBlockDeep(baysTall, totalCount)],
    ['E2-0154', formulas.calculateAntiWalkBlockShallow(baysWide, baysTall, totalCount)],
    ['E2-0611', formulas.calculateSettingBlockIntHorizontal(baysWide, totalCount)],
    ['BE9-2513', formulas.calculateJambFtV(openingHeight, totalCount)],
    ['BE9-2513', formulas.calculateSillFtH(openingWidth, totalCount, baysWide, customBayWidths)],
    ['E9-2512', formulas.calculateFlushFillerV(baysWide, totalCount, openingHeight)],
    ['BE9-2511', formulas.calculateIntVertical(baysWide, totalCount, openingHeight)],
    ['BE9-2515', formulas.calculateOgIntHorizontal(openingWidth, totalCount, baysWide, customBayWidths)],
    ['BE9-2514', formulas.calculateOgHeadH(openingWidth, totalCount, baysWide, customBayWidths)],
    ['BE9-2578', formulas.calculateSillFlashingH(openingWidth, totalCount)],
    ['E9-2519', formulas.calculateGlassStop(openingWidth, baysTall, totalCount, baysWide, customBayWidths)],
    ['E2-0052', formulas.calculateTotalGasketFt(baysWide, baysTall, openingWidth, openingHeight, totalCount)],
  ];

  // --- Build bay dimension arrays ---
  const bayWidths: number[] =
    customBayWidths && customBayWidths.length === baysWide
      ? customBayWidths
      : Array(baysWide).fill(openingWidth / baysWide);

  const bayHeights: number[] =
    customBayHeights && customBayHeights.length === baysTall
      ? customBayHeights
      : Array(baysTall).fill(openingHeight / baysTall);

  // --- Build D.L.O. grid for per-pane glass items ---
  const { dloWidths, dloHeights } = formulas.buildDloGrid(bayWidths, bayHeights);

  // Group panes by unique (DLO width, DLO height) dimensions
  const paneGroups = new Map<string, { width: number; height: number; count: number }>();
  for (let col = 0; col < baysWide; col++) {
    for (let row = 0; row < baysTall; row++) {
      const w = dloWidths[col];
      const h = dloHeights[col][row];
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
  let totalDoorArea = 0.0;
  let totalGlassToAddBack = 0.0;

  if (hasDoors) {
    const doorsWithTotalCount: DoorConfig[] = doors.map((door) => ({
      ...door,
      count: (door.count ?? 0) * totalCount,
    }));
    totalDoorArea = formulas.calculateTotalDoorArea(doorsWithTotalCount);
    totalGlassToAddBack = formulas.calculateGlassToAddBack(doorsWithTotalCount);
  }

  const results: CalculatedOutput[] = [];

  // --- Standard outputs ---
  // Explicitly handle the two BE9-2513 entries to label them correctly
  let be9_2513Counter = 0;

  for (const [partNumber, quantity] of outputs) {
    let desc: string | null = null;
    let partType: string | null = null;
    let resolvedPartNumber = partNumber;

    if (partNumber === 'BE9-2513') {
      be9_2513Counter += 1;
      // Fetch base description
      const baseDesc = PART_NUMBER_MAP['profiles']?.['BE9-2513'] ?? 'UNKNOWN';
      if (be9_2513Counter === 1) {
        desc = `Vertical ${baseDesc} (Jamb)`;
      } else {
        desc = `Horizontal ${baseDesc} (Sill)`;
      }
      partType = 'profiles';
    } else {
      for (const [category, partsDict] of Object.entries(PART_NUMBER_MAP)) {
        if (partNumber in partsDict) {
          const baseDesc = partsDict[partNumber];
          // Add Horizontal/Vertical prefix for specific parts
          if (['BE9-2514', 'BE9-2515', 'E9-2519'].includes(partNumber)) {
            desc = `Horizontal ${baseDesc}`;
          } else if (['E9-2512', 'BE9-2511'].includes(partNumber)) {
            desc = `Vertical ${baseDesc}`;
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

  // --- Per-pane glass outputs ---
  const glassRate = glassPerSqft != null ? Number(glassPerSqft) : 10.5;
  const doorDeduction = totalDoorArea - totalGlassToAddBack;
  const totalPaneArea = Array.from(paneGroups.values()).reduce(
    (sum, g) => sum + (g.width * g.height / 144) * g.count * totalCount, 0,
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
      // Glass make size = DLO + 3/4" per dimension
      const glassMakeW = group.width + 0.75;
      const glassMakeH = group.height + 0.75;
      const paneAreaSqft = (glassMakeW * glassMakeH) / 144;
      const totalPanes = group.count * totalCount;
      const totalAreaSqft = Math.round(paneAreaSqft * totalPanes * 100) / 100;
      glassOutputs.push({
        description: `Glass Pane — DLO: ${group.width.toFixed(2)}" × ${group.height.toFixed(2)}" (${totalPanes} pane${totalPanes !== 1 ? 's' : ''})`,
        quantity: totalAreaSqft,
        part_number: 'N/A',
        type: 'Glass',
        price: glassRate,
        unit: 'sqft',
        pane_count: totalPanes,
        manual: true,
      });
    }

    // Door glass deduction (net of glass-to-add-back)
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

  // Informational door area item
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
