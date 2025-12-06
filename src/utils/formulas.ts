export function calculateRectangleArea(length: number, width: number): number {
  return length * width;
}

export function calculatePerimeter(length: number, width: number): number {
  return 2 * (length + width);
}

export function calculateTotalGasketFt(
  baysWide: number,
  baysTall: number,
  openingWidth: number,
  openingHeight: number,
  totalCount: number
): number {
  const totalInches = (baysWide * 4 * openingHeight) + (baysTall * 4 * openingWidth);
  return (totalInches * totalCount) / 12;
}

export function calculateEndDam(totalCount: number): number {
  return 2 * totalCount;
}

export function calculateWaterDeflector(baysWide: number, totalCount: number): number {
  return 2 * baysWide * totalCount;
}

export function calculateAssemblyScrew(baysWide: number, baysTall: number, totalCount: number): number {
  return ((baysWide * 8) + ((baysTall - 1) * 6 * baysWide)) * totalCount;
}

export function calculateSillFlashScrew(baysWide: number, totalCount: number): number {
  return 3 * baysWide * totalCount;
}

export function calculateEndDamScrew(totalCount: number): number {
  return 4 * totalCount;
}

export function calculateSettingBlockChair(baysWide: number): number {
  return 2 * baysWide;
}

export function calculateSideBlock(baysWide: number, baysTall: number, totalCount: number): number {
  return (baysWide - 1) * baysTall * totalCount;
}

export function calculateSettingBlock(baysWide: number, totalCount: number): number {
  return 2 * baysWide * totalCount;
}

export function calculateAntiWalkBlockDeep(baysTall: number, totalCount: number): number {
  return 2 * baysTall * totalCount;
}

export function calculateAntiWalkBlockShallow(baysWide: number, baysTall: number, totalCount: number): number {
  return (baysWide - 1) * baysTall * totalCount;
}

export function calculateSettingBlockIntHorizontal(baysWide: number, totalCount: number): number {
  return 2 * baysWide * totalCount;
}

export function calculateJambFtV(openingHeight: number, totalCount: number): number[] {
  const singlePieceQty = openingHeight / 12;
  if (totalCount > 1) {
    return Array(totalCount * 2).fill(singlePieceQty);
  }
  return [singlePieceQty, singlePieceQty];
}

export function calculateSillFtH(
  openingWidth: number,
  totalCount: number,
  baysWide?: number,
  customBayWidths?: number[]
): number | number[] {
  if (baysWide && baysWide > 0) {
    let bayWidthsFt: number[];
    if (customBayWidths && customBayWidths.length === baysWide) {
      bayWidthsFt = customBayWidths.map(w => w / 12.0);
    } else {
      const bayWidthFt = (openingWidth / baysWide) / 12;
      bayWidthsFt = Array(baysWide).fill(bayWidthFt);
    }
    if (totalCount > 1) {
      return Array(totalCount).fill(bayWidthsFt).flat();
    }
    return bayWidthsFt;
  }
  const singleInstanceQty = openingWidth / 12;
  if (totalCount > 1) {
    return Array(totalCount).fill(singleInstanceQty);
  }
  return singleInstanceQty;
}

export function calculateFlushFillerV(baysWide: number, totalCount: number, openingHeight: number): number[] {
  const singlePieceQty = openingHeight / 12;
  const piecesPerInstance = baysWide - 1;
  if (totalCount > 1) {
    return Array(piecesPerInstance * totalCount).fill(singlePieceQty);
  }
  return Array(piecesPerInstance).fill(singlePieceQty);
}

export function calculateIntVertical(baysWide: number, totalCount: number, openingHeight: number): number[] {
  const singlePieceQty = openingHeight / 12;
  const piecesPerInstance = baysWide - 1;
  if (totalCount > 1) {
    return Array(piecesPerInstance * totalCount).fill(singlePieceQty);
  }
  return Array(piecesPerInstance).fill(singlePieceQty);
}

export function calculateOgIntHorizontal(
  openingWidth: number,
  totalCount: number,
  baysWide?: number,
  customBayWidths?: number[]
): number | number[] {
  if (baysWide && baysWide > 0) {
    let bayWidthsFt: number[];
    if (customBayWidths && customBayWidths.length === baysWide) {
      bayWidthsFt = customBayWidths.map(w => w / 12.0);
    } else {
      const bayWidthFt = (openingWidth / baysWide) / 12;
      bayWidthsFt = Array(baysWide).fill(bayWidthFt);
    }
    if (totalCount > 1) {
      return Array(totalCount).fill(bayWidthsFt).flat();
    }
    return bayWidthsFt;
  }
  const singleInstanceQty = openingWidth / 12;
  if (totalCount > 1) {
    return Array(totalCount).fill(singleInstanceQty);
  }
  return singleInstanceQty;
}

export function calculateOgHeadH(
  openingWidth: number,
  totalCount: number,
  baysWide?: number,
  customBayWidths?: number[]
): number | number[] {
  if (baysWide && baysWide > 0) {
    let bayWidthsFt: number[];
    if (customBayWidths && customBayWidths.length === baysWide) {
      bayWidthsFt = customBayWidths.map(w => w / 12.0);
    } else {
      const bayWidthFt = (openingWidth / baysWide) / 12;
      bayWidthsFt = Array(baysWide).fill(bayWidthFt);
    }
    if (totalCount > 1) {
      return Array(totalCount).fill(bayWidthsFt).flat();
    }
    return bayWidthsFt;
  }
  const singleInstanceQty = openingWidth / 12;
  if (totalCount > 1) {
    return Array(totalCount).fill(singleInstanceQty);
  }
  return singleInstanceQty;
}

export function calculateSillFlashingH(openingWidth: number, totalCount: number): number | number[] {
  const singleInstanceQty = openingWidth / 12;
  if (totalCount > 1) {
    return Array(totalCount).fill(singleInstanceQty);
  }
  return singleInstanceQty;
}

export function calculateFabricationJoints(baysWide: number, baysTall: number, totalCount: number): number {
  return ((4 * baysWide) + (baysWide * (2 * (baysTall - 1)))) * totalCount;
}

export function calculateGlassStop(
  openingWidth: number,
  baysTall: number,
  totalCount: number,
  baysWide?: number,
  customBayWidths?: number[]
): number | number[] {
  if (baysWide && baysWide > 0) {
    let bayWidthsFt: number[];
    if (customBayWidths && customBayWidths.length === baysWide) {
      bayWidthsFt = customBayWidths.map(w => w / 12.0);
    } else {
      const bayWidthFt = (openingWidth / baysWide) / 12;
      bayWidthsFt = Array(baysWide).fill(bayWidthFt);
    }
    const totalBays = baysWide * baysTall;
    const bayWidthsRepeated = Array(baysTall).fill(bayWidthsFt).flat();
    if (totalCount > 1) {
      return Array(totalCount).fill(bayWidthsRepeated).flat();
    }
    return bayWidthsRepeated;
  }
  const singleInstanceQty = (openingWidth / 12) * baysTall;
  if (totalCount > 1) {
    return Array(totalCount).fill(singleInstanceQty);
  }
  return singleInstanceQty;
}

export function calculateTotalGlass(
  openingWidth: number,
  openingHeight: number,
  totalCount: number,
  baysWide: number,
  baysTall: number
): number {
  return ((openingWidth - (2 * (baysWide + 1))) * (openingHeight - (2 * (baysTall + 1))) * totalCount) / 144;
}

export function calculateDoorSize(doorSizeStr: string): number {
  try {
    const parts = doorSizeStr.toUpperCase().replace(/\s/g, '').split('X');
    if (parts.length !== 2) {
      throw new Error(`Invalid door size format: ${doorSizeStr}`);
    }
    const widthFt = parseFloat(parts[0].replace("'", ""));
    const heightFt = parseFloat(parts[1].replace("'", ""));
    return widthFt * heightFt;
  } catch (e) {
    console.error(`Error calculating door area for '${doorSizeStr}':`, e);
    return 0.0;
  }
}

export function calculateDoorPrice(
  sizeStr: string,
  widthType: string,
  hardwareDict: Record<string, boolean>,
  finish: string
): number {
  const DOOR_PRICES: Record<string, Record<string, Record<string, number>>> = {
    "3x7": {
      "Narrow": { "Clear": 880.00, "Black": 1035.00, "Paint": 1269.00 },
      "Medium": { "Clear": 1180.00, "Black": 1245.00, "Paint": 1653.00 },
      "Wide": { "Clear": 1304.25, "Black": 1413.75, "Paint": 1744.50 }
    },
    "3x8": {
      "Narrow": { "Clear": 921.75, "Black": 1083.00, "Paint": 1328.25 },
      "Medium": { "Clear": 1235.25, "Black": 1304.25, "Paint": 1727.25 },
      "Wide": { "Clear": 1365.00, "Black": 1479.00, "Paint": 1825.50 }
    },
    "3x9": {
      "Narrow": { "Clear": 986.25, "Black": 1159.50, "Paint": 1422.75 },
      "Medium": { "Clear": 1321.50, "Black": 1395.75, "Paint": 1849.50 },
      "Wide": { "Clear": 1461.75, "Black": 1584.00, "Paint": 1953.00 }
    },
    "6x7": {
      "Narrow": { "Clear": 1715.25, "Black": 1863.75, "Paint": 2657.25 },
      "Medium": { "Clear": 2310.75, "Black": 2445.75, "Paint": 3156.75 },
      "Wide": { "Clear": 2559.00, "Black": 2781.00, "Paint": 3435.75 }
    },
    "6x8": {
      "Narrow": { "Clear": 1812.00, "Black": 1970.25, "Paint": 2799.75 },
      "Medium": { "Clear": 2442.00, "Black": 2589.75, "Paint": 3338.25 },
      "Wide": { "Clear": 2700.00, "Black": 2932.50, "Paint": 3630.00 }
    },
    "6x9": {
      "Narrow": { "Clear": 1943.25, "Black": 2111.25, "Paint": 2988.00 },
      "Medium": { "Clear": 2624.25, "Black": 2782.50, "Paint": 3564.00 },
      "Wide": { "Clear": 2901.00, "Black": 3150.00, "Paint": 3861.00 }
    }
  };

  const HARDWARE_PRICES: Record<string, Record<string, number>> = {
    "Concealed Closer": { "Clear": 473.00, "Black": 473.00, "Paint": 473.00 },
    "Exit Devices": { "Clear": 475.00, "Black": 475.00, "Paint": 475.00 },
    "Exit Device": { "Clear": 475.00, "Black": 475.00, "Paint": 475.00 },
    "Continuous Hinges": { "Clear": 285.00, "Black": 375.00, "Paint": 375.00 },
    "Latch Lock w/ Lever Handle": { "Clear": 334.00, "Black": 334.00, "Paint": 334.00 },
    "Lever Handle": { "Clear": 167.00, "Black": 167.00, "Paint": 167.00 },
    "Electric Strike": { "Clear": 355.00, "Black": 355.00, "Paint": 355.00 },
    "Extended Ladder Pull (B2B)": { "Clear": 350.00, "Black": 400.00, "Paint": 400.00 },
    "Extended Ladder Pull (Single)": { "Clear": 175.00, "Black": 200.00, "Paint": 200.00 },
  };

  try {
    const parts = sizeStr.toUpperCase().replace(/'/g, '').replace(/\s/g, '').split('X');
    if (parts.length !== 2) {
      throw new Error(`Invalid door size format: ${sizeStr}`);
    }
    const doorSizeKey = `${parts[0].toLowerCase()}x${parts[1].toLowerCase()}`;

    const finishKey = finish.charAt(0).toUpperCase() + finish.slice(1).toLowerCase();
    const normalizedFinish = ["Clear", "Black", "Paint"].includes(finishKey) ? finishKey : "Clear";

    const basePrice = DOOR_PRICES[doorSizeKey][widthType][normalizedFinish];

    let hardwareTotal = 0;
    for (const [hw, selected] of Object.entries(hardwareDict)) {
      if (selected && HARDWARE_PRICES[hw]) {
        let price = HARDWARE_PRICES[hw][normalizedFinish];
        if ((hw === "Exit Device" || hw === "Exit Devices") && !["3x7", "6x7"].includes(doorSizeKey)) {
          price = 550.00;
        }
        if (doorSizeKey.startsWith("6x")) {
          price *= 2;
        }
        hardwareTotal += price;
      }
    }

    return basePrice + hardwareTotal;
  } catch (e) {
    console.error("Error calculating price:", e);
    return 0.0;
  }
}

export interface DoorInfo {
  size: string;
  count: number;
  stile: string;
  hardware: Record<string, boolean>;
}

export function calculateDoorInfo(doors: DoorInfo[], finish: string = 'Clear'): Array<{
  description: string;
  Style: string;
  quantity: number;
  part_number: string;
  type: string;
  price: number;
  hardware: Record<string, boolean>;
  manual: boolean;
}> {
  const doorItems: Array<{
    description: string;
    Style: string;
    quantity: number;
    part_number: string;
    type: string;
    price: number;
    hardware: Record<string, boolean>;
    manual: boolean;
  }> = [];

  if (doors) {
    for (const doorInfo of doors) {
      const doorSizeStr = doorInfo.size;
      const doorCount = doorInfo.count || 0;
      const doorStile = doorInfo.stile;
      const doorHardware = doorInfo.hardware || {};

      if (doorSizeStr && doorCount > 0) {
        const doorPrice = calculateDoorPrice(doorSizeStr, doorStile, doorHardware, finish);
        doorItems.push({
          description: `Door (${doorSizeStr})`,
          Style: doorStile,
          quantity: doorCount,
          part_number: "N/A",
          type: "Doors",
          price: doorPrice,
          hardware: doorHardware,
          manual: true
        });
      }
    }
  }
  return doorItems;
}

export function calculateTotalDoorArea(doors: DoorInfo[]): number {
  let totalArea = 0.0;
  for (const door of doors) {
    const sizeStr = door.size;
    const count = door.count || 0;
    if (sizeStr && count) {
      const area = calculateDoorSize(sizeStr);
      totalArea += area * count;
    }
  }
  return totalArea;
}

export function calculateGlassToAddBack(doors: DoorInfo[]): number {
  const deductions: Record<string, { height: number; width: number }> = {
    'Narrow': { height: 13.5625, width: 4.875 },
    'Medium': { height: 15.1875, width: 7.625 },
    'Wide': { height: 16.25, width: 10.625 },
  };

  if (!doors || !Array.isArray(doors)) {
    return 0;
  }

  let totalArea = 0.0;

  for (const door of doors) {
    const stile = door.stile?.charAt(0).toUpperCase() + door.stile?.slice(1).toLowerCase() || '';
    if (!deductions[stile]) {
      continue;
    }

    const count = door.count || 1;
    const sizeStr = door.size || '';
    if (!sizeStr) {
      continue;
    }

    const sizeMatch = sizeStr.match(/\s*(\d+)'\s*[xX]\s*(\d+)'/);
    if (!sizeMatch) {
      continue;
    }

    const openingWidthFt = parseInt(sizeMatch[1]);
    const openingHeightFt = parseInt(sizeMatch[2]);

    const openingWidthIn = openingWidthFt * 12;
    const openingHeightIn = openingHeightFt * 12;

    let glassWidth: number;
    if (openingWidthFt === 6) {
      glassWidth = (openingWidthIn / 2) - deductions[stile].width;
    } else {
      glassWidth = openingWidthIn - deductions[stile].width;
    }

    const glassHeight = openingHeightIn - deductions[stile].height;

    if (glassWidth <= 0 || glassHeight <= 0) {
      continue;
    }

    const areaSqft = (glassWidth * glassHeight) / 144;
    totalArea += areaSqft * count;
  }

  return Math.round(totalArea * 100) / 100;
}

