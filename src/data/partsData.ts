// Type definitions and loader for parts data
import partsDataJson from './parts_data.json';

export interface PartData {
  Finish?: string[];
  Length?: string;
  Units?: string;
  "List Price": number | number[];
  "Page Numbers"?: string[];
}

export const partsData: Record<string, PartData> = partsDataJson as Record<string, PartData>;

export default partsData;

