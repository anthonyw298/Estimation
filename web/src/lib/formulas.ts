// TypeScript port of utils/formulas.py
// All formulas, logic, and values are exact copies of the Python source.

export function calculate_rectangle_area(length: number, width: number): number {
  return length * width;
}

export function calculate_perimeter(length: number, width: number): number {
  return 2 * (length + width);
}

export function calculate_total_gasket_ft(
  bays_wide: number,
  bays_tall: number,
  opening_width: number,
  opening_height: number,
  total_count: number
): number {
  const total_inches =
    bays_wide * 4 * opening_height + bays_tall * 4 * opening_width;
  return (total_inches * total_count) / 12;
}

export function calculate_end_dam(total_count: number): number {
  return 2 * total_count;
}

export function calculate_water_deflector(
  bays_wide: number,
  total_count: number
): number {
  return 2 * bays_wide * total_count;
}

export function calculate_assembly_screw(
  bays_wide: number,
  bays_tall: number,
  total_count: number
): number {
  return (bays_wide * 8 + (bays_tall - 1) * 6 * bays_wide) * total_count;
}

export function calculate_sill_flash_screw(
  bays_wide: number,
  total_count: number
): number {
  return 3 * bays_wide * total_count;
}

export function calculate_end_dam_screw(total_count: number): number {
  return 4 * total_count;
}

export function calculate_setting_block_chair(
  bays_wide: number,
  total_count: number
): number {
  return 2 * bays_wide * total_count;
}

export function calculate_side_block(
  bays_wide: number,
  bays_tall: number,
  total_count: number
): number {
  return (bays_wide - 1) * bays_tall * total_count;
}

export function calculate_setting_block(
  bays_wide: number,
  total_count: number
): number {
  return 2 * bays_wide * total_count;
}

export function calculate_anti_walk_block_deep(
  bays_tall: number,
  total_count: number
): number {
  return 2 * bays_tall * total_count;
}

export function calculate_anti_walk_block_shallow(
  bays_wide: number,
  bays_tall: number,
  total_count: number
): number {
  return (bays_wide - 1) * bays_tall * total_count;
}

export function calculate_setting_block_int_horizontal(
  bays_wide: number,
  total_count: number
): number {
  return 2 * bays_wide * total_count;
}

export function calculate_jamb_ft_v(
  opening_height: number,
  total_count: number
): number[] {
  /**
   * Calculates vertical jamb feet. Returns a list of full height (split into 2 pieces - left and right).
   * Associated with profile: BE9-2513 (jamb uses full height, 2 pieces)
   */
  const single_piece_qty = opening_height / 12; // Convert inches to feet
  // Always return as list with 2 pieces per instance (left and right)
  const pair = [single_piece_qty, single_piece_qty];
  if (total_count > 1) {
    const result: number[] = [];
    for (let i = 0; i < total_count; i++) {
      result.push(...pair);
    }
    return result;
  }
  return pair;
}

export function calculate_sill_ft_h(
  opening_width: number,
  total_count: number,
  bays_wide: number | null = null,
  custom_bay_widths: number[] | null = null
): number | number[] {
  /**
   * Calculates horizontal sill feet. Returns a list of bay widths if bays_wide is provided,
   * otherwise returns the whole width. Returns a list if total_count > 1, else a float.
   * Uses custom_bay_widths if provided, otherwise divides equally.
   * Associated with profile: BE9-2513 (sill only uses bay widths)
   */
  if (bays_wide && bays_wide > 0) {
    let bay_widths_ft: number[];
    // Use custom bay widths if provided, otherwise divide equally
    if (custom_bay_widths && custom_bay_widths.length === bays_wide) {
      // Convert custom bay widths from inches to feet
      bay_widths_ft = custom_bay_widths.map((w) => w / 12.0);
    } else {
      // Equal division
      const bay_width_ft = opening_width / bays_wide / 12;
      bay_widths_ft = Array(bays_wide).fill(bay_width_ft);
    }

    if (total_count > 1) {
      const result: number[] = [];
      for (let i = 0; i < total_count; i++) {
        result.push(...bay_widths_ft);
      }
      return result;
    }
    return bay_widths_ft;
  }

  const single_instance_qty = opening_width / 12;
  if (total_count > 1) {
    return Array(total_count).fill(single_instance_qty);
  }
  return single_instance_qty;
}

export function calculate_flush_filler_v(
  bays_wide: number,
  total_count: number,
  opening_height: number
): number[] {
  /**
   * Calculates vertical flush filler feet. Returns a list of full height for each filler.
   * Associated with profile: E9-2512 (uses full height, one piece per filler)
   */
  const single_piece_qty = opening_height / 12; // Convert inches to feet
  // Number of fillers = (bays_wide - 1), one piece each
  const pieces_per_instance = bays_wide - 1;
  if (total_count > 1) {
    return Array(pieces_per_instance * total_count).fill(single_piece_qty);
  }
  return Array(pieces_per_instance).fill(single_piece_qty);
}

export function calculate_int_vertical(
  bays_wide: number,
  total_count: number,
  opening_height: number
): number[] {
  /**
   * Calculates intermediate vertical feet. Returns a list of full height for each mullion.
   * Associated with profile: BE9-2511 (uses full height, one piece per mullion)
   */
  const single_piece_qty = opening_height / 12; // Convert inches to feet
  // Number of mullions = (bays_wide - 1), one piece each
  const pieces_per_instance = bays_wide - 1;
  if (total_count > 1) {
    return Array(pieces_per_instance * total_count).fill(single_piece_qty);
  }
  return Array(pieces_per_instance).fill(single_piece_qty);
}

export function calculate_og_int_horizontal(
  opening_width: number,
  total_count: number,
  bays_wide: number | null = null,
  custom_bay_widths: number[] | null = null
): number | number[] {
  /**
   * Calculates outside glazing intermediate horizontal feet. Returns a list of bay widths if bays_wide is provided,
   * otherwise returns the whole width. Uses custom_bay_widths if provided, otherwise divides equally.
   * Returns a list if total_count > 1, else a float.
   * Associated with profile: BE9-2515 (uses bay widths)
   */
  if (bays_wide && bays_wide > 0) {
    let bay_widths_ft: number[];
    if (custom_bay_widths && custom_bay_widths.length === bays_wide) {
      bay_widths_ft = custom_bay_widths.map((w) => w / 12.0);
    } else {
      const bay_width_ft = opening_width / bays_wide / 12;
      bay_widths_ft = Array(bays_wide).fill(bay_width_ft);
    }

    if (total_count > 1) {
      const result: number[] = [];
      for (let i = 0; i < total_count; i++) {
        result.push(...bay_widths_ft);
      }
      return result;
    }
    return bay_widths_ft;
  }

  const single_instance_qty = opening_width / 12;
  if (total_count > 1) {
    return Array(total_count).fill(single_instance_qty);
  }
  return single_instance_qty;
}

export function calculate_og_head_h(
  opening_width: number,
  total_count: number,
  bays_wide: number | null = null,
  custom_bay_widths: number[] | null = null
): number | number[] {
  /**
   * Calculates outside glazing head horizontal feet. Returns a list of bay widths if bays_wide is provided,
   * otherwise returns the whole width. Uses custom_bay_widths if provided, otherwise divides equally.
   * Returns a list if total_count > 1, else a float.
   * Associated with profile: BE9-2514 (uses bay widths)
   */
  if (bays_wide && bays_wide > 0) {
    let bay_widths_ft: number[];
    if (custom_bay_widths && custom_bay_widths.length === bays_wide) {
      bay_widths_ft = custom_bay_widths.map((w) => w / 12.0);
    } else {
      const bay_width_ft = opening_width / bays_wide / 12;
      bay_widths_ft = Array(bays_wide).fill(bay_width_ft);
    }

    if (total_count > 1) {
      const result: number[] = [];
      for (let i = 0; i < total_count; i++) {
        result.push(...bay_widths_ft);
      }
      return result;
    }
    return bay_widths_ft;
  }

  const single_instance_qty = opening_width / 12;
  if (total_count > 1) {
    return Array(total_count).fill(single_instance_qty);
  }
  return single_instance_qty;
}

export function calculate_sill_flashing_h(
  opening_width: number,
  total_count: number
): number | number[] {
  /**
   * Calculates sill flashing horizontal feet. Returns a list if total_count > 1, else a float.
   * Associated with profile: BE9-2578
   */
  const single_instance_qty = opening_width / 12;
  if (total_count > 1) {
    return Array(total_count).fill(single_instance_qty);
  }
  return single_instance_qty;
}

export function calculate_fabrication_joints(
  bays_wide: number,
  bays_tall: number,
  total_count: number
): number {
  return (4 * bays_wide + bays_wide * (2 * (bays_tall - 1))) * total_count;
}

export function calculate_glass_stop(
  opening_width: number,
  bays_tall: number,
  total_count: number,
  bays_wide: number | null = null,
  custom_bay_widths: number[] | null = null
): number | number[] {
  /**
   * Calculate glass stop length. Returns a list of bay widths repeated for each bay (total bays = bays_wide * bays_tall).
   * Uses custom_bay_widths if provided, otherwise divides equally.
   * Returns a list if total_count > 1, else a float.
   * Associated with profile: E9-2519 (uses bay widths, one per bay)
   */
  if (bays_wide && bays_wide > 0) {
    let bay_widths_ft: number[];
    if (custom_bay_widths && custom_bay_widths.length === bays_wide) {
      bay_widths_ft = custom_bay_widths.map((w) => w / 12.0);
    } else {
      const bay_width_ft = opening_width / bays_wide / 12;
      bay_widths_ft = Array(bays_wide).fill(bay_width_ft);
    }

    // Total number of bays = bays_wide * bays_tall
    // Each bay gets one glass stop piece, so repeat bay_widths_ft for each row (bays_tall)
    const bay_widths_repeated: number[] = [];
    for (let i = 0; i < bays_tall; i++) {
      bay_widths_repeated.push(...bay_widths_ft);
    }

    if (total_count > 1) {
      const result: number[] = [];
      for (let i = 0; i < total_count; i++) {
        result.push(...bay_widths_repeated);
      }
      return result;
    }
    return bay_widths_repeated;
  }

  const single_instance_qty = (opening_width / 12) * bays_tall;
  if (total_count > 1) {
    return Array(total_count).fill(single_instance_qty);
  }
  return single_instance_qty;
}

export function calculate_total_glass(
  opening_width: number,
  opening_height: number,
  total_count: number,
  bays_wide: number,
  bays_tall: number
): number {
  return (
    ((opening_width - 2 * (bays_wide + 1)) *
      (opening_height - 2 * (bays_tall + 1)) *
      total_count) /
    144
  );
}

export function calculate_door_size(door_size_str: string): number {
  /**
   * Calculates the area of a single door from a size string like "3' X 7'".
   * Returns the area in square feet.
   */
  try {
    // Normalize the string and split on 'X'
    const parts = door_size_str.toUpperCase().replace(/ /g, "").split("X");
    if (parts.length !== 2) {
      throw new Error(`Invalid door size format: ${door_size_str}`);
    }

    // Remove the apostrophe and convert to float
    const width_ft = parseFloat(parts[0].replace(/'/g, ""));
    const height_ft = parseFloat(parts[1].replace(/'/g, ""));

    const area = width_ft * height_ft;
    return area;
  } catch (e) {
    console.error(
      `Error calculating door area for '${door_size_str}': ${e}`
    );
    return 0.0;
  }
}

// Door price matrix
const DOOR_PRICES: Record<string, Record<string, Record<string, number>>> = {
  "3x7": {
    Narrow: { Clear: 880.0, Black: 1035.0, Paint: 1269.0 },
    Medium: { Clear: 1180.0, Black: 1245.0, Paint: 1653.0 },
    Wide: { Clear: 1304.25, Black: 1413.75, Paint: 1744.5 },
  },
  "3x8": {
    Narrow: { Clear: 921.75, Black: 1083.0, Paint: 1328.25 },
    Medium: { Clear: 1235.25, Black: 1304.25, Paint: 1727.25 },
    Wide: { Clear: 1365.0, Black: 1479.0, Paint: 1825.5 },
  },
  "3x9": {
    Narrow: { Clear: 986.25, Black: 1159.5, Paint: 1422.75 },
    Medium: { Clear: 1321.5, Black: 1395.75, Paint: 1849.5 },
    Wide: { Clear: 1461.75, Black: 1584.0, Paint: 1953.0 },
  },
  "6x7": {
    Narrow: { Clear: 1715.25, Black: 1863.75, Paint: 2657.25 },
    Medium: { Clear: 2310.75, Black: 2445.75, Paint: 3156.75 },
    Wide: { Clear: 2559.0, Black: 2781.0, Paint: 3435.75 },
  },
  "6x8": {
    Narrow: { Clear: 1812.0, Black: 1970.25, Paint: 2799.75 },
    Medium: { Clear: 2442.0, Black: 2589.75, Paint: 3338.25 },
    Wide: { Clear: 2700.0, Black: 2932.5, Paint: 3630.0 },
  },
  "6x9": {
    Narrow: { Clear: 1943.25, Black: 2111.25, Paint: 2988.0 },
    Medium: { Clear: 2624.25, Black: 2782.5, Paint: 3564.0 },
    Wide: { Clear: 2901.0, Black: 3150.0, Paint: 3861.0 },
  },
};

// Hardware base prices by finish
const HARDWARE_PRICES: Record<string, Record<string, number>> = {
  "Concealed Closer": { Clear: 473.0, Black: 473.0, Paint: 473.0 },
  "Exit Devices": { Clear: 475.0, Black: 475.0, Paint: 475.0 }, // "Exit Devices" in UI
  "Exit Device": { Clear: 475.0, Black: 475.0, Paint: 475.0 }, // Legacy key
  "Continuous Hinges": { Clear: 285.0, Black: 375.0, Paint: 375.0 },
  "Latch Lock w/ Lever Handle": { Clear: 334.0, Black: 334.0, Paint: 334.0 }, // "Latch Lock w/ Lever Handle" in UI
  "Lever Handle": { Clear: 167.0, Black: 167.0, Paint: 167.0 }, // Estimation: Half of Latch+Lever set
  "Electric Strike": { Clear: 355.0, Black: 355.0, Paint: 355.0 },
  "Extended Ladder Pull (B2B)": { Clear: 350.0, Black: 400.0, Paint: 400.0 }, // Placeholder price
  "Extended Ladder Pull (Single)": {
    Clear: 175.0,
    Black: 200.0,
    Paint: 200.0,
  }, // Placeholder price
};

export function calculate_door_price(
  size_str: string,
  width_type: string,
  hardware_dict: Record<string, boolean>,
  finish: string
): number {
  /**
   * Calculates total price of a door given size (e.g. "3' X 8'"), width_type ("Narrow", "Medium", "Wide"),
   * finish ("Clear", "Black", "Paint"), and selected hardware.
   * Hardware prices also vary by finish.
   */
  try {
    // Normalize size key (e.g., "3' X 8'" -> "3x8")
    const parts = size_str
      .toUpperCase()
      .replace(/'/g, "")
      .replace(/ /g, "")
      .split("X");
    if (parts.length !== 2) {
      throw new Error(`Invalid door size format: ${size_str}`);
    }
    const door_size_key = `${parts[0].toLowerCase()}x${parts[1].toLowerCase()}`;

    // Normalize finish to Title Case (e.g. "clear" -> "Clear") to match keys
    let finish_key =
      finish.charAt(0).toUpperCase() + finish.slice(1).toLowerCase();
    if (!["Clear", "Black", "Paint"].includes(finish_key)) {
      finish_key = "Clear"; // Default to Clear if unknown
    }

    // Get base price
    const base_price =
      DOOR_PRICES[door_size_key]?.[width_type]?.[finish_key];
    if (base_price === undefined) {
      throw new Error(
        `Invalid key in pricing lookup: ${door_size_key}, ${width_type}, ${finish_key}`
      );
    }

    // Calculate hardware cost based on finish
    let hardware_total = 0;
    for (const [hw, selected] of Object.entries(hardware_dict)) {
      if (selected && hw in HARDWARE_PRICES) {
        let price = HARDWARE_PRICES[hw][finish_key];

        // Special rule: Exit Device is $550 for all doors except 3x7 and 6x7
        if (
          (hw === "Exit Device" || hw === "Exit Devices") &&
          !["3x7", "6x7"].includes(door_size_key)
        ) {
          price = 550.0;
        }

        // Double price for double doors (all hardware)
        if (door_size_key.startsWith("6x")) {
          price *= 2;
        }

        hardware_total += price;
      }
    }

    return base_price + hardware_total;
  } catch (e) {
    console.error(`Error calculating price: ${e}`);
    return 0.0;
  }
}

interface DoorInput {
  size?: string;
  count?: number;
  stile?: string;
  hardware?: Record<string, boolean>;
}

interface DoorItem {
  description: string;
  Style: string | undefined;
  quantity: number;
  part_number: string;
  type: string;
  price: number;
  hardware: Record<string, boolean>;
  manual: boolean;
}

export function calculate_door_info(
  doors: DoorInput[],
  finish: string = "Clear",
  total_count: number = 1
): DoorItem[] {
  /**
   * Takes a list of door inputs and returns a list of dictionaries with door information,
   * including calculated price and other details.
   *
   * Note: door count is per elevation, so it's multiplied by total_count to get total quantity.
   */
  const door_items: DoorItem[] = [];
  if (doors) {
    for (const door_info of doors) {
      const door_size_str = door_info.size;
      const door_count_per_elev = door_info.count ?? 0; // Count is per elevation
      const door_stile = door_info.stile;
      const door_hardware = door_info.hardware ?? {};

      if (door_size_str && door_count_per_elev > 0) {
        // Calculate total door count (per elevation * total_count)
        const total_door_count = door_count_per_elev * total_count;
        const door_price = calculate_door_price(
          door_size_str,
          door_stile!,
          door_hardware,
          finish
        );

        door_items.push({
          description: `Door (${door_size_str})`,
          Style: door_stile,
          quantity: total_door_count, // Total quantity across all elevations
          part_number: "N/A",
          type: "Doors",
          price: door_price,
          hardware: door_hardware,
          manual: true,
        });
      }
    }
  }
  return door_items;
}

export function calculate_total_door_area(doors: DoorInput[]): number {
  /**
   * Calculates the total area (in sqft) of all doors in the list.
   */
  let total_area = 0.0;
  for (const door of doors) {
    const size_str = door.size;
    const count = door.count ?? 0;
    if (size_str && count) {
      const area = calculate_door_size(size_str);
      total_area += area * count;
    }
  }
  return total_area;
}

interface DoorWithStile {
  size?: string;
  count?: number;
  stile?: string;
}

export function calculate_glass_to_add_back(
  doors: DoorWithStile[] | null | undefined
): number {
  /**
   * Calculate total glass back area in sqft based on door sizes.
   */
  const deductions: Record<string, { height: number; width: number }> = {
    Narrow: { height: 13.5625, width: 4.875 },
    Medium: { height: 15.1875, width: 7.625 },
    Wide: { height: 16.25, width: 10.625 },
  };

  if (!doors || !Array.isArray(doors)) {
    return 0;
  }

  let total_area = 0.0;

  for (const door of doors) {
    const raw_stile = door.stile ?? "";
    // Title case: capitalize first letter, lowercase the rest
    const stile =
      raw_stile.charAt(0).toUpperCase() + raw_stile.slice(1).toLowerCase();
    if (!(stile in deductions)) {
      continue;
    }

    const count = door.count ?? 1;
    const size_str = door.size ?? "";
    if (!size_str) {
      continue;
    }

    // Parse door opening width and height in feet from 'size' string
    const size_match = size_str.match(/^\s*(\d+)' *[xX] *(\d+)'/);
    if (!size_match) {
      continue;
    }

    const opening_width_ft = parseInt(size_match[1], 10);
    const opening_height_ft = parseInt(size_match[2], 10);

    // Convert to inches
    const opening_width_in = opening_width_ft * 12;
    const opening_height_in = opening_height_ft * 12;

    // For 6' width doors, width deduction applies after dividing width by 2 (paired door)
    let glass_width: number;
    if (opening_width_ft === 6) {
      glass_width = opening_width_in / 2 - deductions[stile].width;
    } else {
      glass_width = opening_width_in - deductions[stile].width;
    }

    const glass_height = opening_height_in - deductions[stile].height;

    if (glass_width <= 0 || glass_height <= 0) {
      continue;
    }

    const area_sqft = (glass_width * glass_height) / 144;
    total_area += area_sqft * count;
  }

  return Math.round(total_area * 100) / 100;
}

// camelCase aliases so yes45tu.ts can use `formulas.calculateEndDam(...)` etc.
export const calculateRectangleArea = calculate_rectangle_area;
export const calculatePerimeter = calculate_perimeter;
export const calculateTotalGasketFt = calculate_total_gasket_ft;
export const calculateEndDam = calculate_end_dam;
export const calculateWaterDeflector = calculate_water_deflector;
export const calculateAssemblyScrew = calculate_assembly_screw;
export const calculateSillFlashScrew = calculate_sill_flash_screw;
export const calculateEndDamScrew = calculate_end_dam_screw;
export const calculateSettingBlockChair = calculate_setting_block_chair;
export const calculateSideBlock = calculate_side_block;
export const calculateSettingBlock = calculate_setting_block;
export const calculateAntiWalkBlockDeep = calculate_anti_walk_block_deep;
export const calculateAntiWalkBlockShallow = calculate_anti_walk_block_shallow;
export const calculateSettingBlockIntHorizontal = calculate_setting_block_int_horizontal;
export const calculateJambFtV = calculate_jamb_ft_v;
export const calculateSillFtH = calculate_sill_ft_h;
export const calculateFlushFillerV = calculate_flush_filler_v;
export const calculateIntVertical = calculate_int_vertical;
export const calculateOgIntHorizontal = calculate_og_int_horizontal;
export const calculateOgHeadH = calculate_og_head_h;
export const calculateSillFlashingH = calculate_sill_flashing_h;
export const calculateFabricationJoints = calculate_fabrication_joints;
export const calculateGlassStop = calculate_glass_stop;
export const calculateTotalGlass = calculate_total_glass;
export const calculateDoorSize = calculate_door_size;
export const calculateDoorPrice = calculate_door_price;
export const calculateDoorInfo = calculate_door_info;
export const calculateTotalDoorArea = calculate_total_door_area;
export const calculateGlassToAddBack = calculate_glass_to_add_back;
