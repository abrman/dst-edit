export interface SheetData {
  [key: string]: string
}

export interface ParsedSheet {
  id: string
  name: string
  number: string
  properties: Record<string, string>
}
