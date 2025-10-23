"use client";

import { useState } from "react";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import type { SheetData } from "@/lib/types";

interface DataTableProps {
  sheets: SheetData[];
  onUpdate: (sheets: SheetData[]) => void;
}

export function DataTable({ sheets, onUpdate }: DataTableProps) {
  const [editingCell, setEditingCell] = useState<{
    row: number;
    col: string;
  } | null>(null);

  if (sheets.length === 0) return null;

  const columns = Object.keys(sheets[0]).filter((col) => col !== "ID" && col !== "Name");

  const handleCellChange = (rowIndex: number, column: string, value: string) => {
    const newSheets = [...sheets];
    newSheets[rowIndex] = {
      ...newSheets[rowIndex],
      [column]: value,
    };

    if (column === "Number" || column === "Title") {
      const number = column === "Number" ? value : newSheets[rowIndex]["Number"] || "";
      const title = column === "Title" ? value : newSheets[rowIndex]["Title"] || "";
      newSheets[rowIndex]["Name"] = `${number} ${title}`.trim();
    }

    onUpdate(newSheets);
  };

  return (
    <div className="border rounded-lg overflow-hidden">
      <div className="overflow-x-auto">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead className="w-12 bg-muted">#</TableHead>
              {columns.map((col) => (
                <TableHead key={col} className="bg-muted font-semibold">
                  {col}
                </TableHead>
              ))}
            </TableRow>
          </TableHeader>
          <TableBody>
            {sheets.map((sheet, rowIndex) => (
              <TableRow key={rowIndex}>
                <TableCell className="font-medium text-muted-foreground">{rowIndex + 1}</TableCell>
                {columns.map((col) => (
                  <TableCell
                    key={`${rowIndex}-${col}`}
                    className="p-0"
                    onClick={() => setEditingCell({ row: rowIndex, col })}
                  >
                    <Input
                      value={sheet[col] || ""}
                      onChange={(e) => handleCellChange(rowIndex, col, e.target.value)}
                      onBlur={() => setEditingCell(null)}
                      className={
                        editingCell?.row === rowIndex && editingCell?.col === col
                          ? "border-primary"
                          : "border-transparent"
                      }
                    />
                  </TableCell>
                ))}
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
    </div>
  );
}
