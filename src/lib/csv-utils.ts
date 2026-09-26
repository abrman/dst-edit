import type { DstModel, SheetNode, TreeNode } from "./dst-model"
import { allSheets, mapSheets } from "./tree"
import { cleanSheetValue } from "./autocad-text"

// Columns that describe a sheet but can't be changed through a CSV
const READ_ONLY_COLUMNS = ["ID", "Folder", "Layout", "Drawing"]

export function editableColumns(model: DstModel): string[] {
  return [...model.sheetColumns, ...model.customProps.filter((d) => d.scope === "sheet").map((d) => d.name)]
}

function sheetsWithFolders(nodes: TreeNode[], path: string[] = []): { sheet: SheetNode; folder: string }[] {
  return nodes.flatMap((n) =>
    n.kind === "sheet" ? [{ sheet: n, folder: path.join(" / ") }] : sheetsWithFolders(n.children, [...path, n.name]),
  )
}

function csvField(value: string): string {
  return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value
}

export function exportToCSV(model: DstModel): string {
  const columns = editableColumns(model)
  const header = ["ID", "Folder", ...columns, "Drawing"]
  const rows = sheetsWithFolders(model.children).map(({ sheet, folder }) => [
    sheet.id,
    folder,
    ...columns.map((c) => sheet.props[c] ?? ""),
    sheet.layout.fileName || sheet.layout.relative,
  ])
  // BOM so Excel reads the file as UTF-8, CRLF because that's what Excel writes
  return "﻿" + [header, ...rows].map((row) => row.map(csvField).join(",")).join("\r\n")
}

/** Reads CSV text into rows, handling quoted commas, quotes and line breaks, and any line ending. */
export function parseCSV(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ""
  let inQuotes = false
  text = text.replace(/^﻿/, "")

  for (let i = 0; i < text.length; i++) {
    const char = text[i]
    if (inQuotes) {
      if (char === '"' && text[i + 1] === '"') {
        field += '"'
        i++
      } else if (char === '"') {
        inQuotes = false
      } else {
        field += char
      }
    } else if (char === '"') {
      inQuotes = true
    } else if (char === ",") {
      row.push(field)
      field = ""
    } else if (char === "\r" || char === "\n") {
      if (char === "\r" && text[i + 1] === "\n") i++
      row.push(field)
      rows.push(row)
      row = []
      field = ""
    } else {
      field += char
    }
  }
  if (field !== "" || row.length) {
    row.push(field)
    rows.push(row)
  }
  return rows.filter((r) => r.some((f) => f.trim() !== ""))
}

/** Decodes a CSV file as UTF-8, falling back to Windows-1252 for files Excel saved as plain "CSV". */
export function decodeCSVFile(buffer: ArrayBuffer): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(buffer)
  } catch {
    return new TextDecoder("windows-1252").decode(buffer)
  }
}

export interface CSVImportResult {
  model: DstModel
  updatedSheets: number
  unmatchedRows: number
  newColumns: string[]
  /** Values where "/" was replaced, since AutoCAD doesn't allow it in sheet titles */
  replacedSlashes: number
}

/** Applies CSV rows to the sheets, matching rows by ID (or by position when there is no ID column). */
export function importCSV(model: DstModel, rows: string[][]): CSVImportResult {
  if (rows.length < 2) throw new Error("The CSV has no rows to import.")
  const header = rows[0].map((h) => h.trim())
  const body = rows.slice(1)
  const sheets = allSheets(model.children)

  const idIndex = header.indexOf("ID")
  if (idIndex < 0 && body.length !== sheets.length) {
    throw new Error(
      `The CSV has no ID column and ${body.length} rows, but the sheet set has ${sheets.length} sheets, so rows can't be matched to sheets.`,
    )
  }

  const editable = new Set(editableColumns(model))
  const columns = header
    .map((name, index) => ({ name, index }))
    .filter(({ name }) => name && !READ_ONLY_COLUMNS.includes(name))
  const newColumns = columns.map((c) => c.name).filter((name) => !editable.has(name))

  const byId = new Map<string, string[]>()
  let unmatchedRows = 0
  const knownIds = new Set(sheets.map((s) => s.id))
  body.forEach((row, i) => {
    const id = idIndex >= 0 ? (row[idIndex] ?? "").trim() : sheets[i].id
    if (knownIds.has(id)) byId.set(id, row)
    else unmatchedRows++
  })

  let updatedSheets = 0
  let replacedSlashes = 0
  const children = mapSheets(model.children, (sheet) => {
    const row = byId.get(sheet.id)
    if (!row) return sheet
    const props = { ...sheet.props }
    for (const { name, index } of columns) {
      const raw = row[index] ?? ""
      const value = cleanSheetValue(name, raw)
      if (value !== raw) replacedSlashes++
      // Don't create empty values for properties the sheet doesn't have
      if (!(name in props) && value === "") continue
      props[name] = value
    }
    const changed = Object.keys(props).some((k) => props[k] !== sheet.props[k])
    if (!changed) return sheet
    updatedSheets++
    return { ...sheet, props }
  })

  return {
    model: {
      ...model,
      children,
      customProps: [...model.customProps, ...newColumns.map((name) => ({ name, scope: "sheet" as const, value: "" }))],
    },
    updatedSheets,
    unmatchedRows,
    newColumns,
    replacedSlashes,
  }
}
