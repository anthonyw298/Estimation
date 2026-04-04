import { PART_NUMBER_MAP } from '@/data/part-number';
import * as formulas from '@/lib/formulas';
import { CalculatedOutput, DoorConfig } from '@/types';

/**
 * Calculates all the specific output quantities for the 'YES 45TU Center Set' system
 * by calling dedicated formula functions.
 * Returns a list of objects with description, quantity, part number, and type.
 *
 * Key differences from Front Set (OG):
 * - Different profile part numbers (BE9-2551 jamb, BE9-2553 head, BE9-2556 horizontal, etc.)
 * - Head is 1 piece full width (not per bay)
 * - Sill is separate part (BE9-2579) at bay widths
 * - Different accessory formulas for water deflector, assembly screws, anti-walk, setting blocks
 * - New shear block components (E1-1058, E1-1059) with associated screws
 * - No setting block chair, sill setting block, anti-walk deep, or side blocks
 */
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
  // Safety check for doors
  if (!doors) {
    doors = [];
  }

  // --- Center set accessory formulas ---
  const waterDeflector = 2 * (baysTall - 1) * baysWide * totalCount;
  const assemblyScrewPC1216 = 4 * (baysWide - 1) * (baysTall + 1) * totalCount;
  const assemblyScrewPC1220 = 4 * (1 + baysWide * baysTall) * totalCount;
  const antiWalkBlock = baysWide * baysTall * totalCount; // 1 per lite
  const settingBlockOutside = 2 * baysWide * baysTall * totalCount;
  const settingBlockInside = 2 * baysWide * baysTall * totalCount;
  const flatFiller = (2 * baysWide + 4) * totalCount;

  // --- Shear block components (center set only) ---
  const shearBlockSillHoriz = 2 * baysWide * baysTall * totalCount;
  const shearBlockHead = 2 * baysWide * totalCount;
  const shearBlockScrewPC1028 = 2 * shearBlockSillHoriz;
  const shearBlockScrewFC1212 = shearBlockHead;
  const shearBlockScrewPC1210 = shearBlockSillHoriz;

  const outputs: [string, number | number[]][] = [
    // --- Accessories ---
    ['E1-0199', formulas.calculateEndDam(totalCount)],
    ['E2-0047', waterDeflector],
    ['PC-1216', assemblyScrewPC1216],
    ['PC-1220', assemblyScrewPC1220],
    ['PM-1008-SS', formulas.calculateSillFlashScrew(baysWide, totalCount)],
    ['UA-1212', formulas.calculateEndDamScrew(totalCount)],
    ['E2-0153', antiWalkBlock],
    ['E2-0020', settingBlockOutside],
    ['E2-0611', settingBlockInside],
    ['E1-1054', flatFiller],
    // --- Shear blocks ---
    ['E1-1058', shearBlockSillHoriz],
    ['E1-1059', shearBlockHead],
    ['PC-1028', shearBlockScrewPC1028],
    ['FC-1212', shearBlockScrewFC1212],
    ['PC-1210', shearBlockScrewPC1210],
    // --- Profiles ---
    ['BE9-2551', formulas.calculateJambFtV(openingHeight, totalCount)],
    ['BE9-2579', formulas.calculateSillFtH(openingWidth, totalCount, baysWide, customBayWidths)],
    ['BE9-2552', formulas.calculateFlushFillerV(baysWide, totalCount, openingHeight)],
    ['BE9-2555', formulas.calculateIntVertical(baysWide, totalCount, openingHeight)],
    ['BE9-2556', formulas.calculateOgIntHorizontal(openingWidth, totalCount, baysWide, customBayWidths)],
    ['BE9-2553', formulas.calculateSillFlashingH(openingWidth, totalCount)], // Head: 1 piece full width
    ['BE9-2578', formulas.calculateSillFlashingH(openingWidth, totalCount)],
    ['E9-1015', formulas.calculateGlassStop(openingWidth, baysTall, totalCount, baysWide, customBayWidths)],
    ['E2-0052', formulas.calculateTotalGasketFt(baysWide, baysTall, openingWidth, openingHeight, totalCount)],
  ];

  // --- Total area calculations (D.L.O. based) ---
  const totalGlassArea = formulas.calculateTotalGlass(
    openingWidth, openingHeight, totalCount, baysWide, baysTall, customBayWidths, customBayHeights,
  );

  // Only calculate door area if doors exist
  const hasDoors = doors && doors.length > 0;
  let totalDoorArea = 0.0;
  let totalGlassToAddBack = 0.0;
  let adjustedGlassArea: number;

  if (hasDoors) {
    const doorsWithTotalCount: DoorConfig[] = doors.map((door) => ({
      ...door,
      count: (door.count ?? 0) * totalCount,
    }));
    totalDoorArea = formulas.calculateTotalDoorArea(doorsWithTotalCount);
    totalGlassToAddBack = formulas.calculateGlassToAddBack(doorsWithTotalCount);
    adjustedGlassArea = Math.max(totalGlassArea - totalDoorArea + totalGlassToAddBack, 0);
  } else {
    totalDoorArea = 0.0;
    totalGlassToAddBack = 0.0;
    adjustedGlassArea = totalGlassArea;
  }

  const results: CalculatedOutput[] = [];

  // --- Standard outputs ---
  for (const [partNumber, quantity] of outputs) {
    let desc: string | null = null;
    let partType: string | null = null;
    let resolvedPartNumber = partNumber;

    for (const [category, partsDict] of Object.entries(PART_NUMBER_MAP)) {
      if (partNumber in partsDict) {
        const baseDesc = partsDict[partNumber];
        // Add Horizontal/Vertical prefix for specific parts
        if (['BE9-2553', 'BE9-2556', 'E9-1015'].includes(partNumber)) {
          desc = `Horizontal ${baseDesc}`;
        } else if (['BE9-2552', 'BE9-2555'].includes(partNumber)) {
          desc = `Vertical ${baseDesc}`;
        } else if (partNumber === 'BE9-2551') {
          desc = `Vertical ${baseDesc} (Jamb)`;
        } else if (partNumber === 'BE9-2579') {
          desc = `Horizontal ${baseDesc} (Sill)`;
        } else {
          desc = baseDesc;
        }
        partType = category;
        break;
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

  // Check if the adjusted glass area is zero and add a specific message.
  let glassOutput: CalculatedOutput;
  if (adjustedGlassArea === 0) {
    glassOutput = {
      description: 'Glass Area (Adjusted)',
      quantity: 0,
      part_number: 'N/A',
      type: 'Glass',
      price: 0.0,
      unit: 'sqft',
      manual: true,
      message: 'Total door area equals or exceeds total glass area. No glass is needed.',
    };
  } else {
    glassOutput = {
      description: 'Glass Area (Adjusted)',
      quantity: adjustedGlassArea,
      part_number: 'N/A',
      type: 'Glass',
      price: glassPerSqft != null ? Number(glassPerSqft) : 10.5,
      unit: 'sqft',
      manual: true,
    };
  }

  const fabPrice = fabricationCostPerJoint != null ? Number(fabricationCostPerJoint) : 15.0;
  const manualOutputs: CalculatedOutput[] = [
    glassOutput,
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

  // Only include door area calculation if doors exist
  if (hasDoors) {
    manualOutputs.splice(1, 0, {
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
