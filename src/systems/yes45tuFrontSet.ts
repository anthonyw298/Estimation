import { PART_NUMBER_MAP } from '../data/partNumber';
import {
  calculateTotalGasketFt,
  calculateEndDam,
  calculateWaterDeflector,
  calculateAssemblyScrew,
  calculateSillFlashScrew,
  calculateEndDamScrew,
  calculateSettingBlockChair,
  calculateSideBlock,
  calculateSettingBlock,
  calculateAntiWalkBlockDeep,
  calculateAntiWalkBlockShallow,
  calculateSettingBlockIntHorizontal,
  calculateJambFtV,
  calculateSillFtH,
  calculateFlushFillerV,
  calculateIntVertical,
  calculateOgIntHorizontal,
  calculateOgHeadH,
  calculateSillFlashingH,
  calculateGlassStop,
  calculateTotalGlass,
  calculateFabricationJoints,
  calculateTotalDoorArea,
  calculateGlassToAddBack,
  DoorInfo
} from '../utils/formulas';

export interface CalculatedOutput {
  description: string;
  quantity: number | number[];
  part_number: string;
  type: string;
  price?: number;
  unit?: string;
  manual?: boolean;
  message?: string;
}

export function calculateYes45tuQuantities(
  baysWide: number,
  baysTall: number,
  totalCount: number,
  openingWidth: number,
  openingHeight: number,
  doors: DoorInfo[] = [],
  customBayWidths?: number[]
): CalculatedOutput[] {
  const outputs: Array<[string, number | number[]]> = [
    ["E1-0199", calculateEndDam(totalCount)],
    ["E2-0047", calculateWaterDeflector(baysWide, totalCount)],
    ["PC-1220", calculateAssemblyScrew(baysWide, baysTall, totalCount)],
    ["PM-1006-SS", calculateSillFlashScrew(baysWide, totalCount)],
    ["UA-1212", calculateEndDamScrew(totalCount)],
    ["E1-2530", calculateSettingBlockChair(baysWide)],
    ["E2-0166", calculateSideBlock(baysWide, baysTall, totalCount)],
    ["E2-0177", calculateSettingBlock(baysWide, totalCount)],
    ["E2-0545", calculateAntiWalkBlockDeep(baysTall, totalCount)],
    ["E2-0154", calculateAntiWalkBlockShallow(baysWide, baysTall, totalCount)],
    ["E2-0611", calculateSettingBlockIntHorizontal(baysWide, totalCount)],
    ["BE9-2513", calculateJambFtV(openingHeight, totalCount)],
    ["BE9-2513", calculateSillFtH(openingWidth, totalCount, baysWide, customBayWidths)],
    ["E9-2512", calculateFlushFillerV(baysWide, totalCount, openingHeight)],
    ["BE9-2511", calculateIntVertical(baysWide, totalCount, openingHeight)],
    ["BE9-2515", calculateOgIntHorizontal(openingWidth, totalCount, baysWide, customBayWidths)],
    ["BE9-2514", calculateOgHeadH(openingWidth, totalCount, baysWide, customBayWidths)],
    ["BE9-2578", calculateSillFlashingH(openingWidth, totalCount)],
    ["E9-2519", calculateGlassStop(openingWidth, baysTall, totalCount, baysWide, customBayWidths)],
    ["E2-0052", calculateTotalGasketFt(baysWide, baysTall, openingWidth, openingHeight, totalCount)]
  ];

  const results: CalculatedOutput[] = [];
  let be9_2513_counter = 0;

  for (const [partNumber, quantity] of outputs) {
    let desc: string | null = null;
    let partType: string | null = null;

    if (partNumber === "BE9-2513") {
      be9_2513_counter++;
      const baseDesc = PART_NUMBER_MAP.profiles[partNumber] || "UNKNOWN";
      if (be9_2513_counter === 1) {
        desc = `Vertical ${baseDesc} (Jamb)`;
      } else {
        desc = `Horizontal ${baseDesc} (Sill)`;
      }
      partType = "profiles";
    } else {
      for (const [category, partsDict] of Object.entries(PART_NUMBER_MAP)) {
        if (partNumber in partsDict) {
          const baseDesc = (partsDict as Record<string, string>)[partNumber];
          if (["BE9-2514", "BE9-2515", "E9-2519"].includes(partNumber)) {
            desc = `Horizontal ${baseDesc}`;
          } else if (["E9-2512", "BE9-2511"].includes(partNumber)) {
            desc = `Vertical ${baseDesc}`;
          } else {
            desc = baseDesc;
          }
          partType = category;
          break;
        }
      }
    }

    if (!desc) {
      desc = "UNKNOWN";
      partType = "UNKNOWN";
    }

    results.push({
      description: desc,
      quantity: quantity,
      part_number: partNumber,
      type: partType || "UNKNOWN"
    });
  }

  const totalGlassArea = calculateTotalGlass(openingWidth, openingHeight, totalCount, baysWide, baysTall);
  const totalDoorArea = calculateTotalDoorArea(doors);
  const totalGlassToAddBack = calculateGlassToAddBack(doors);
  const adjustedGlassArea = Math.max(totalGlassArea - totalDoorArea + totalGlassToAddBack, 0);

  if (adjustedGlassArea === 0) {
    results.push({
      description: "Glass Area (Adjusted)",
      quantity: 0,
      part_number: "N/A",
      type: "Glass",
      price: 0.0,
      unit: 'sqft',
      manual: true,
      message: "Total door area equals or exceeds total glass area. No glass is needed."
    });
  } else {
    results.push({
      description: "Glass Area (Adjusted)",
      quantity: adjustedGlassArea,
      part_number: "N/A",
      type: "Glass",
      price: 10.5,
      unit: 'sqft',
      manual: true
    });
  }

  results.push({
    description: "Door Area (to subtract from glass)",
    quantity: totalDoorArea,
    part_number: "N/A",
    type: "Calculations",
    unit: 'sqft',
    manual: true
  });

  results.push({
    description: "Joints Fabrication Labor",
    quantity: calculateFabricationJoints(baysWide, baysTall, totalCount),
    part_number: "N/A",
    type: "Fabrication",
    price: 15.0,
    unit: 'joints',
    manual: true
  });

  return results;
}

