// lib/types.ts
export interface SheetData {
  [key: string]: string;
}

export interface ParsedSheet {
  id: string;
  name: string;
  number: string;
  properties: Record<string, string>;
}

export interface PropertyMetadata {
  isCustom: boolean;
  vt: string;
}

export interface ParseResult {
  sheets: SheetData[];
  propertyMetadata: Map<string, PropertyMetadata>;
}
