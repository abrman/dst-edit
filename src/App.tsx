"use client";

import { useState } from "react";
import { FileUploader } from "@/components/file-uploader";
import { DataTable } from "@/components/data-table";
import { ExportButtons } from "@/components/export-buttons";
import { parseDSTFile, reconstructXML } from "@/lib/dst-parser";
import { decodeDST, encodeDST } from "@/lib/bit-flip";
import type { SheetData } from "@/lib/types";

export default function DSTEditorPage() {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [originalXML, setOriginalXML] = useState<string>("");
  const [fileName, setFileName] = useState<string>("");

  const handleFileUpload = async (file: File) => {
    const arrayBuffer = await file.arrayBuffer();
    const xmlText = decodeDST(arrayBuffer);

    setOriginalXML(xmlText);
    setFileName(file.name);

    const parsedSheets = parseDSTFile(xmlText);
    setSheets(parsedSheets);
  };

  const handleCSVImport = (csvData: SheetData[]) => {
    // Validate dimensions match
    if (csvData.length === sheets.length && csvData[0] && sheets[0]) {
      const csvKeys = Object.keys(csvData[0]);
      const sheetKeys = Object.keys(sheets[0]);

      if (csvKeys.length === sheetKeys.length) {
        setSheets(csvData);
      } else {
        alert("CSV dimensions do not match the original table");
      }
    } else {
      alert("CSV dimensions do not match the original table");
    }
  };

  const handleDSTExport = () => {
    const xml = reconstructXML(originalXML, sheets);
    const encodedBuffer = encodeDST(xml);
    const blob = new Blob([encodedBuffer], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName || "edited-sheet-set.dst";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b">
        <div className="container mx-auto px-4 py-6">
          <h1 className="text-3xl font-bold">DST File Editor</h1>
          <p className="text-muted-foreground mt-2">Edit AutoCAD Sheet Set files in a table format</p>
        </div>
      </header>

      <main className="container mx-auto px-4 py-8">
        {sheets.length === 0 ? (
          <FileUploader onFileUpload={handleFileUpload} />
        ) : (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-xl font-semibold">{fileName}</h2>
                <p className="text-sm text-muted-foreground">
                  {sheets.length} sheet{sheets.length !== 1 ? "s" : ""} loaded
                </p>
              </div>
              <ExportButtons sheets={sheets} onDSTExport={handleDSTExport} onCSVImport={handleCSVImport} />
            </div>

            <DataTable sheets={sheets} onUpdate={setSheets} />
          </div>
        )}
      </main>
    </div>
  );
}
