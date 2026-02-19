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

  // --- Total area calculations ---
  const totalGlassArea = formulas.calculateTotalGlass(openingWidth, openingHeight, totalCount, baysWide, baysTall);

  // Only calculate door area if doors exist
  // Door count is per elevation, so multiply by totalCount for total calculations
  const hasDoors = doors && doors.length > 0;
  let totalDoorArea = 0.0;
  let totalGlassToAddBack = 0.0;
  let adjustedGlassArea: number;

  if (hasDoors) {
    // Multiply door counts by totalCount since door count is per elevation
    const doorsWithTotalCount: DoorConfig[] = doors.map((door) => ({
      ...door,
      count: (door.count ?? 0) * totalCount,
    }));
    totalDoorArea = formulas.calculateTotalDoorArea(doorsWithTotalCount);
    totalGlassToAddBack = formulas.calculateGlassToAddBack(doorsWithTotalCount);
    adjustedGlassArea = Math.max(totalGlassArea - totalDoorArea + totalGlassToAddBack, 0); // Prevent negative glass area
  } else {
    totalDoorArea = 0.0;
    totalGlassToAddBack = 0.0;
    adjustedGlassArea = totalGlassArea; // No door adjustments needed
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
