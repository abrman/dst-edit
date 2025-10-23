import type { SheetData } from "./types"

export function exportToCSV(sheets: SheetData[]): string {
  if (sheets.length === 0) return ""

  const headers = Object.keys(sheets[0])
  const rows = sheets.map((sheet) =>
    headers.map((header) => {
      const value = sheet[header] || ""
      // Escape quotes and wrap in quotes if contains comma, quote, or newline
      if (value.includes(",") || value.includes('"') || value.includes("\n")) {
        return `"${value.replace(/"/g, '""')}"`
      }
      return value
    }),
  )

  const csvContent = [headers.join(","), ...rows.map((row) => row.join(","))].join("\n")

  return csvContent
}

export function parseCSV(csvString: string): SheetData[] {
  const lines = csvString.split("\n").filter((line) => line.trim())
  if (lines.length === 0) return []

  const headers = parseCSVLine(lines[0])
  const sheets: SheetData[] = []

  for (let i = 1; i < lines.length; i++) {
    const values = parseCSVLine(lines[i])
    const sheet: SheetData = {}

    headers.forEach((header, index) => {
      sheet[header] = values[index] || ""
    })

    sheets.push(sheet)
  }

  return sheets
}

function parseCSVLine(line: string): string[] {
  const result: string[] = []
  let current = ""
  let inQuotes = false

  for (let i = 0; i < line.length; i++) {
    const char = line[i]
    const nextChar = line[i + 1]

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        current += '"'
        i++ // Skip next quote
      } else {
        inQuotes = !inQuotes
      }
    } else if (char === "," && !inQuotes) {
      result.push(current)
      current = ""
    } else {
      current += char
    }
  }

  result.push(current)
  return result
}
