"use client";

import type React from "react";

import { useRef } from "react";
import { Download, Upload, FileCode } from "lucide-react";
import { Button } from "@/components/ui/button";
import { exportToCSV, parseCSV } from "@/lib/csv-utils";
import type { SheetData } from "@/lib/types";
import { Tooltip, TooltipTrigger, TooltipContent, TooltipProvider } from "@/components/ui/tooltip";

interface ExportButtonsProps {
  sheets: SheetData[];
  onDSTExport: () => void;
  onCSVImport: (data: SheetData[]) => void;
}

export function ExportButtons({ sheets, onDSTExport, onCSVImport }: ExportButtonsProps) {
  const csvInputRef = useRef<HTMLInputElement>(null);

  const handleCSVExport = () => {
    const csv = exportToCSV(sheets);
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "sheet-data.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleCSVImportClick = () => {
    csvInputRef.current?.click();
  };

  const handleCSVFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0];
      if (file) {
        const text = await file.text();
        const data = parseCSV(text);
        onCSVImport(data);
      }
    }
  };

  return (
    <TooltipProvider>
      <div className="flex gap-2">
        {/* Export CSV with tooltip */}
        <Tooltip>
          <TooltipTrigger asChild>
            <Button onClick={handleCSVExport} variant="outline">
              <Download className="mr-2 h-4 w-4" />
              Export CSV
            </Button>
          </TooltipTrigger>
          <TooltipContent>
            <p>Download as CSV for editing within Excel or alike</p>
          </TooltipContent>
        </Tooltip>

        {/* Import CSV with tooltip */}
        <Tooltip>
          <TooltipTrigger asChild>
            <Button onClick={handleCSVImportClick} variant="outline">
              <Upload className="mr-2 h-4 w-4" />
              Import CSV
            </Button>
          </TooltipTrigger>
          <TooltipContent>
            <p>Upload updated CSV with your changes</p>
          </TooltipContent>
        </Tooltip>

        <input ref={csvInputRef} type="file" accept=".csv" onChange={handleCSVFileChange} className="hidden" />

        {/* DST Export – tooltip optional */}
        <Tooltip>
          <TooltipTrigger asChild>
            <Button onClick={onDSTExport}>
              <FileCode className="mr-2 h-4 w-4" />
              Export DST
            </Button>
          </TooltipTrigger>
          <TooltipContent>
            <p>Export DST file for use withing AutoCAD</p>
          </TooltipContent>
        </Tooltip>
      </div>
    </TooltipProvider>
  );
}
