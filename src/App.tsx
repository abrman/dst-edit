import { useEffect, useMemo, useRef, useState } from "react";
import type React from "react";
import { Download, FileCode, FolderOpen, Redo2, Undo2, Upload } from "lucide-react";
import { toast } from "sonner";
import { FileUploader } from "@/components/file-uploader";
import { SheetsView } from "@/components/sheets-view";
import { SheetSetSettings } from "@/components/sheet-set-settings";
import { CommitInput } from "@/components/commit-input";
import { Button } from "@/components/ui/button";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { Toaster } from "@/components/ui/sonner";
import { useHistory } from "@/hooks/use-history";
import { decodeDST, encodeDST } from "@/lib/bit-flip";
import { applyModel, countBrokenProps, parseDst, type DstModel } from "@/lib/dst-model";
import { decodeCSVFile, exportToCSV, importCSV, parseCSV } from "@/lib/csv-utils";
import { allFolders, allSheets } from "@/lib/tree";
import { inferBaseDir } from "@/lib/paths";
import { isTitleTooLong, TITLE_MAX_LENGTH } from "@/lib/autocad-text";

interface LoadedFile {
  name: string;
  xml: string;
  model: DstModel;
}

function download(data: BlobPart, fileName: string, type: string) {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

function isTextInput(target: EventTarget | null) {
  return target instanceof HTMLElement && (target.tagName === "INPUT" || target.tagName === "TEXTAREA");
}

function IconButton({ label, children, ...props }: React.ComponentProps<typeof Button> & { label: string }) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <Button variant="ghost" size="icon" aria-label={label} {...props}>
          {children}
        </Button>
      </TooltipTrigger>
      <TooltipContent>{label}</TooltipContent>
    </Tooltip>
  );
}

export default function App() {
  const [file, setFile] = useState<LoadedFile | null>(null);
  const history = useHistory<DstModel | null>(null);
  const model = history.value;
  const [saved, setSaved] = useState<DstModel | null>(null);
  const csvInput = useRef<HTMLInputElement>(null);
  const dstInput = useRef<HTMLInputElement>(null);

  const dirty = model !== null && model !== saved;
  const original = useMemo(
    () => new Map(file ? allSheets(file.model.children).map((s) => [s.id, s]) : []),
    [file],
  );
  const baseDir = useMemo(() => (file ? inferBaseDir(allSheets(file.model.children)) : null), [file]);

  const onChange = (update: (m: DstModel) => DstModel) => history.set((m) => (m ? update(m) : m));

  const openFile = async (f: File) => {
    if (dirty && !window.confirm("Open another file? Your unsaved changes to this one will be lost.")) return;
    try {
      const xml = decodeDST(await f.arrayBuffer());
      const parsed = parseDst(xml);
      setFile({ name: f.name, xml, model: parsed });
      history.reset(parsed);
      setSaved(parsed);
      const broken = countBrokenProps(xml);
      if (broken) {
        toast.info(`This file has ${broken} damaged duplicate properties`, {
          description:
            "They were added by an older version of DST Edit when importing a CSV. They're hidden here and will be removed when you save.",
          duration: 10000,
        });
      }
    } catch (error) {
      toast.error("Couldn't open that file", { description: (error as Error).message });
    }
  };

  const saveDST = () => {
    if (!file || !model) return;
    const tooLong = allSheets(model.children).filter((s) => isTitleTooLong(s.props.Title));
    if (tooLong.length) {
      toast.error(`${tooLong.length} sheet title(s) are longer than ${TITLE_MAX_LENGTH} characters`, {
        description: `AutoCAD can't use them, so shorten them before saving: ${tooLong
          .slice(0, 3)
          .map((s) => s.props.Number || s.props.Title.slice(0, 20) + "…")
          .join(", ")}${tooLong.length > 3 ? "…" : ""}. They're marked in red.`,
      });
      return;
    }
    try {
      download(encodeDST(applyModel(file.xml, model)), file.name, "application/octet-stream");
      setSaved(model);
    } catch (error) {
      toast.error("Couldn't save the DST", { description: (error as Error).message });
    }
  };

  const exportCSV = () => {
    if (!file || !model) return;
    download(exportToCSV(model), file.name.replace(/\.dst$/i, "") + ".csv", "text/csv");
  };

  const importCSVFile = async (f: File) => {
    if (!model) return;
    try {
      const result = importCSV(model, parseCSV(decodeCSVFile(await f.arrayBuffer())));
      history.set(result.model);
      const notes = [
        result.unmatchedRows && `${result.unmatchedRows} row(s) didn't match a sheet and were skipped.`,
        result.newColumns.length && `New custom properties: ${result.newColumns.join(", ")}.`,
        result.replacedSlashes &&
          `“/” was replaced with “ ̸” in ${result.replacedSlashes} value(s), since AutoCAD doesn't allow it.`,
      ].filter(Boolean);
      toast.success(`Updated ${result.updatedSheets} sheet(s) from ${f.name}`, {
        description: notes.join(" ") || undefined,
      });
    } catch (error) {
      toast.error("Couldn't import the CSV", { description: (error as Error).message });
    }
  };

  // Undo/redo outside text fields (inside them the browser undoes typing), and Ctrl+S to save
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      const key = e.key.toLowerCase();
      if (key === "s") {
        e.preventDefault();
        (document.activeElement as HTMLElement | null)?.blur(); // commit the cell being edited first
        setTimeout(saveDST);
      } else if (!isTextInput(e.target) && key === "z") {
        e.preventDefault();
        if (e.shiftKey) history.redo();
        else history.undo();
      } else if (!isTextInput(e.target) && key === "y") {
        e.preventDefault();
        history.redo();
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  });

  useEffect(() => {
    if (!dirty) return;
    const warn = (e: BeforeUnloadEvent) => e.preventDefault();
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [dirty]);

  const sheetCount = model ? allSheets(model.children).length : 0;
  const folderCount = model ? allFolders(model.children).length : 0;

  return (
    <TooltipProvider>
      <div className="min-h-screen bg-background">
        <header className="sticky top-0 z-20 border-b bg-background/95 backdrop-blur">
          <div className="container mx-auto flex flex-wrap items-center gap-x-3 gap-y-2 px-4 py-2.5">
            <span className="font-bold tracking-tight">DST Edit</span>
            {file && (
              <>
                <span className="text-muted-foreground/50">/</span>
                <span className="max-w-64 truncate text-sm text-muted-foreground" title={file.name}>
                  {file.name}
                </span>
                {dirty && (
                  <span className="flex items-center gap-1.5 text-xs text-amber-700 dark:text-amber-400">
                    <span className="size-1.5 rounded-full bg-current" /> Unsaved changes
                  </span>
                )}
                <div className="ml-auto flex items-center gap-1">
                  <IconButton label="Undo (Ctrl+Z)" onClick={history.undo} disabled={!history.canUndo}>
                    <Undo2 />
                  </IconButton>
                  <IconButton label="Redo (Ctrl+Y)" onClick={history.redo} disabled={!history.canRedo}>
                    <Redo2 />
                  </IconButton>
                  <IconButton label="Open another DST" onClick={() => dstInput.current?.click()}>
                    <FolderOpen />
                  </IconButton>
                  <div className="mx-1 h-6 w-px bg-border" />
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button variant="outline" onClick={() => csvInput.current?.click()}>
                        <Download /> Import CSV
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent>Bring in a CSV edited in Excel. Rows are matched to sheets by ID.</TooltipContent>
                  </Tooltip>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button variant="outline" onClick={exportCSV}>
                        <Upload /> Export CSV
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent>Download the sheet list for editing in Excel</TooltipContent>
                  </Tooltip>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button onClick={saveDST}>
                        <FileCode /> Save DST
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent>Download the edited sheet set for AutoCAD (Ctrl+S)</TooltipContent>
                  </Tooltip>
                </div>
              </>
            )}
          </div>
        </header>

        <input
          ref={csvInput}
          type="file"
          accept=".csv,text/csv"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            e.target.value = "";
            if (f) importCSVFile(f);
          }}
        />
        <input
          ref={dstInput}
          type="file"
          accept=".dst"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            e.target.value = "";
            if (f) openFile(f);
          }}
        />

        <main className="container mx-auto px-4 py-6">
          {!model ? (
            <div className="mx-auto max-w-2xl space-y-6 py-10">
              <div>
                <h1 className="text-3xl font-bold">Sheet set editor</h1>
                <p className="mt-2 text-muted-foreground">
                  Open an AutoCAD sheet set (.dst) to rename, renumber, reorder and organize its sheets. Everything
                  happens in your browser; the file isn't uploaded anywhere.
                </p>
              </div>
              <FileUploader onFileUpload={openFile} />
            </div>
          ) : (
            <div className="space-y-5">
              <div>
                <CommitInput
                  value={model.setProps.Name ?? ""}
                  onCommit={(Name) => onChange((m) => ({ ...m, setProps: { ...m.setProps, Name } }))}
                  placeholder="Untitled sheet set"
                  aria-label="Sheet set name"
                  title="Sheet set name, as shown in AutoCAD"
                  className="-ml-2 h-11 max-w-3xl text-2xl font-bold"
                />
                <p className="mt-1 text-sm text-muted-foreground">
                  {sheetCount} sheet{sheetCount === 1 ? "" : "s"} in {folderCount} folder{folderCount === 1 ? "" : "s"}
                  {model.setProps.ProjectName && model.setProps.ProjectName !== model.setProps.Name
                    ? ` · ${model.setProps.ProjectName}`
                    : ""}
                </p>
              </div>

              <Tabs defaultValue="sheets">
                <TabsList>
                  <TabsTrigger value="sheets">Sheets</TabsTrigger>
                  <TabsTrigger value="settings">Sheet set</TabsTrigger>
                </TabsList>
                {/* Kept mounted so selection and collapsed folders survive switching tabs */}
                <TabsContent value="sheets" forceMount className="mt-3 data-[state=inactive]:hidden">
                  <SheetsView model={model} original={original} baseDir={baseDir} onChange={onChange} />
                </TabsContent>
                <TabsContent value="settings" className="mt-3">
                  <SheetSetSettings model={model} baseDir={baseDir} onChange={onChange} />
                </TabsContent>
              </Tabs>
            </div>
          )}
        </main>
        <Toaster position="bottom-right" />
      </div>
    </TooltipProvider>
  );
}
