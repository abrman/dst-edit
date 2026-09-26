import { useMemo, useState } from "react";
import {
  ChevronsDownUp,
  ChevronsUpDown,
  FolderInput,
  FolderPlus,
  ListOrdered,
  Search,
  Trash2,
  TriangleAlert,
  X,
} from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Kbd } from "@/components/ui/kbd";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import { TreeGrid } from "@/components/tree-grid";
import { BulkEditDialog } from "@/components/bulk-edit-dialog";
import { MoveDialog } from "@/components/move-dialog";
import { generateID, type DstModel, type FolderNode, type SheetNode, type TreeNode } from "@/lib/dst-model";
import {
  allFolders,
  allSheets,
  deleteNodes,
  findNode,
  mapSheets,
  moveNodes,
  orderedIds,
  shiftNodes,
  targetSheets,
  updateNode,
  visibleRows,
  type DropPosition,
} from "@/lib/tree";
import { editableColumns } from "@/lib/csv-utils";
import { cleanSheetValue } from "@/lib/autocad-text";
import { isUserSpecific } from "@/lib/paths";
import { DrawingPathDialog } from "@/components/drawing-path-dialog";

interface SheetsViewProps {
  model: DstModel;
  original: Map<string, SheetNode>;
  /** The folder the .dst is in, when it can be worked out; used for relative drawing paths */
  baseDir: string | null;
  onChange: (update: (model: DstModel) => DstModel) => void;
}

const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "s"}`;

export function SheetsView({ model, original, baseDir, onChange }: SheetsViewProps) {
  const [selection, setSelection] = useState<Set<string>>(new Set());
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  const [filter, setFilter] = useState("");
  const [focusFolderId, setFocusFolderId] = useState<string | null>(null);
  const [bulkOpen, setBulkOpen] = useState(false);
  const [bulkColumn, setBulkColumn] = useState("Number");
  const [moveIds, setMoveIds] = useState<string[] | null>(null);
  const [deleteIds, setDeleteIds] = useState<string[] | null>(null);
  const [pathSheetId, setPathSheetId] = useState<string | null>(null);

  const columns = useMemo(() => editableColumns(model), [model]);
  const customColumns = useMemo(
    () => new Set(model.customProps.filter((d) => d.scope === "sheet").map((d) => d.name)),
    [model.customProps],
  );
  const rows = useMemo(() => visibleRows(model.children, collapsed, filter), [model.children, collapsed, filter]);
  const targets = useMemo(() => targetSheets(model.children, selection), [model.children, selection]);
  const sheets = useMemo(() => allSheets(model.children), [model.children]);
  const userSpecificCount = sheets.filter((s) => isUserSpecific(s.layout)).length;
  const pathSheet = pathSheetId ? (sheets.find((s) => s.id === pathSheetId) ?? null) : null;
  const updateChildren = (fn: (children: TreeNode[]) => TreeNode[]) =>
    onChange((m) => {
      const children = fn(m.children);
      return children === m.children ? m : { ...m, children };
    });

  const editSheet = (id: string, column: string, value: string) =>
    updateChildren((children) =>
      updateNode(children, id, (n) => (n.kind === "sheet" ? { ...n, props: { ...n.props, [column]: value } } : n)),
    );

  const pasteGrid = (sheetId: string, column: string, grid: string[][]) => {
    const sheetRows = rows.flatMap((r) => (r.node.kind === "sheet" ? [r.node] : []));
    const rowStart = sheetRows.findIndex((s) => s.id === sheetId);
    const colStart = columns.indexOf(column);
    const edits = new Map<string, Record<string, string>>();
    let cells = 0;
    let replaced = 0;
    grid.forEach((line, r) => {
      const sheet = sheetRows[rowStart + r];
      if (!sheet) return;
      line.forEach((raw, c) => {
        const col = columns[colStart + c];
        if (!col) return;
        const value = cleanSheetValue(col, raw);
        if (value !== raw) replaced++;
        edits.set(sheet.id, { ...edits.get(sheet.id), [col]: value });
        cells++;
      });
    });
    updateChildren((children) =>
      mapSheets(children, (s) => (edits.has(s.id) ? { ...s, props: { ...s.props, ...edits.get(s.id) } } : s)),
    );
    const skipped = grid.reduce((n, line) => n + line.length, 0) - cells;
    const notes = [
      skipped && `${skipped} didn't fit and were left out.`,
      replaced && `“/” was replaced with “ ̸” in ${plural(replaced, "cell")}, since AutoCAD doesn't allow it.`,
    ].filter(Boolean);
    toast.success(`Pasted ${plural(cells, "cell")}`, { description: notes.join(" ") || undefined });
  };

  const editFolder = (id: string, patch: Partial<Pick<FolderNode, "name" | "desc">>) =>
    updateChildren((children) => updateNode(children, id, (n) => (n.kind === "folder" ? { ...n, ...patch } : n)));

  const move = (ids: string[], targetId: string | null, position: DropPosition) => {
    updateChildren((children) => moveNodes(children, ids, targetId, position));
    if (targetId && position === "inside") setCollapsed((c) => new Set([...c].filter((id) => id !== targetId)));
  };

  const newFolder = (parentId?: string) => {
    const folder: FolderNode = { kind: "folder", id: generateID(), name: "New folder", desc: "", children: [] };
    const selected = parentId ? [] : orderedIds(model.children, selection);
    updateChildren((children) => {
      if (parentId) {
        return updateNode(children, parentId, (n) =>
          n.kind === "folder" ? { ...n, children: [...n.children, folder] } : n,
        );
      }
      if (selected.length === 0) return [...children, folder];
      // Group the selection: the folder goes where the first selected row is, and the selection moves into it
      const placed = moveNodes([...children, folder], [folder.id], selected[0], "before");
      return moveNodes(placed, selected, folder.id, "inside");
    });
    if (parentId) setCollapsed((c) => new Set([...c].filter((id) => id !== parentId)));
    setSelection(new Set());
    setFilter("");
    setFocusFolderId(folder.id);
  };

  const confirmDelete = () => {
    if (!deleteIds) return;
    const ids = new Set(deleteIds);
    updateChildren((children) => deleteNodes(children, ids));
    setSelection((s) => new Set([...s].filter((id) => !ids.has(id))));
    setDeleteIds(null);
  };

  const deleteSummary = useMemo(() => {
    const nodes = (deleteIds ?? []).map((id) => findNode(model.children, id)).filter(Boolean) as TreeNode[];
    return {
      sheets: nodes.filter((n) => n.kind === "sheet").length,
      folders: nodes.filter((n) => n.kind === "folder").length,
    };
  }, [deleteIds, model.children]);

  const selectedSheets = targets.length;
  const folderCount = allFolders(model.children).length;

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2">
        <div className="relative w-full sm:w-72">
          <Search className="pointer-events-none absolute top-1/2 left-2.5 size-4 -translate-y-1/2 text-muted-foreground" />
          <Input
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
            placeholder="Filter sheets and folders"
            className="pl-8"
          />
        </div>

        <Button variant="outline" onClick={() => newFolder()}>
          <FolderPlus />
          {selection.size ? "New folder from selection" : "New folder"}
        </Button>
        <Button
          variant="outline"
          onClick={() => {
            setBulkOpen(true);
          }}
        >
          <ListOrdered />
          Edit column…
        </Button>

        {selection.size > 0 && (
          <div className="flex items-center gap-1 rounded-md bg-primary/6 py-0.5 pr-0.5 pl-3">
            <span className="mr-1 text-sm font-medium">{selection.size} selected</span>
            <Button variant="ghost" size="sm" onClick={() => setMoveIds([...selection])}>
              <FolderInput /> Move to…
            </Button>
            <Button
              variant="ghost"
              size="sm"
              className="text-destructive hover:text-destructive"
              onClick={() => setDeleteIds([...selection])}
            >
              <Trash2 /> Remove
            </Button>
            <Button variant="ghost" size="icon-sm" onClick={() => setSelection(new Set())} aria-label="Clear selection">
              <X />
            </Button>
          </div>
        )}

        {userSpecificCount > 0 && (
          <button
            type="button"
            onClick={() => setPathSheetId(sheets.find((s) => isUserSpecific(s.layout))?.id ?? null)}
            className="flex items-center gap-1.5 rounded-md bg-orange-50 px-2.5 py-1.5 text-sm text-orange-800 hover:bg-orange-100 dark:bg-orange-500/10 dark:text-orange-300"
            title="These drawings are stored in one user's folder, so other users can't open them"
          >
            <TriangleAlert className="size-4" />
            {plural(userSpecificCount, "sheet")} only work{userSpecificCount === 1 ? "s" : ""} for one user
          </button>
        )}

        <div className="ml-auto flex items-center gap-1">
          <Button
            variant="ghost"
            size="icon-sm"
            title="Expand all folders"
            onClick={() => setCollapsed(new Set())}
            disabled={folderCount === 0}
          >
            <ChevronsUpDown />
          </Button>
          <Button
            variant="ghost"
            size="icon-sm"
            title="Collapse all folders"
            onClick={() => setCollapsed(new Set(allFolders(model.children).map((f) => f.id)))}
            disabled={folderCount === 0}
          >
            <ChevronsDownUp />
          </Button>
        </div>
      </div>

      {rows.length === 0 ? (
        <div className="rounded-lg border border-dashed p-10 text-center text-sm text-muted-foreground">
          {filter ? `Nothing matches “${filter}”.` : "This sheet set has no sheets."}
        </div>
      ) : (
        <TreeGrid
          rows={rows}
          columns={columns}
          customColumns={customColumns}
          original={original}
          selection={selection}
          onSelectionChange={setSelection}
          collapsed={collapsed}
          onToggleCollapsed={(id) =>
            setCollapsed((c) => {
              const next = new Set(c);
              if (!next.delete(id)) next.add(id);
              return next;
            })
          }
          dragEnabled={!filter}
          focusFolderId={focusFolderId}
          onFolderFocused={() => setFocusFolderId(null)}
          onEditSheet={editSheet}
          onPasteGrid={pasteGrid}
          onEditFolder={editFolder}
          onMove={move}
          onShift={(ids, direction) => updateChildren((children) => shiftNodes(children, new Set(ids), direction))}
          onMoveTo={setMoveIds}
          onDelete={setDeleteIds}
          onNewFolderInside={(id) => newFolder(id)}
          onModifyDrawingPath={setPathSheetId}
        />
      )}

      <p className="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-muted-foreground">
        <span>
          Paste from Excel into any cell to fill several rows at once.
        </span>
        <span>
          <Kbd>Enter</Kbd> / <Kbd>↑</Kbd> <Kbd>↓</Kbd> move between rows
        </span>
        <span>
          <Kbd>Alt</Kbd>+<Kbd>↑</Kbd> <Kbd>↓</Kbd> or drag the handle to reorder
        </span>
        <span>
          <Kbd>Shift</Kbd>+click to select a range
        </span>
        <span>Changed cells are highlighted.</span>
        <span>“/” in numbers, titles and descriptions becomes “ ̸”, since AutoCAD doesn't allow it.</span>
      </p>

      <BulkEditDialog
        open={bulkOpen}
        onOpenChange={setBulkOpen}
        tree={model.children}
        targets={targets}
        columns={columns}
        initialColumn={columns.includes(bulkColumn) ? bulkColumn : columns[0]}
        scopeLabel={selection.size ? `the ${plural(selectedSheets, "selected sheet")}` : `all ${plural(selectedSheets, "sheet")}`}
        onApply={(column, changes) => {
          setBulkColumn(column);
          updateChildren((children) =>
            mapSheets(children, (s) => (changes.has(s.id) ? { ...s, props: { ...s.props, [column]: changes.get(s.id)! } } : s)),
          );
        }}
      />

      <DrawingPathDialog
        sheet={pathSheet}
        sheets={sheets}
        baseDir={baseDir}
        onOpenChange={(open) => !open && setPathSheetId(null)}
        onApply={(layouts) => {
          updateChildren((children) =>
            mapSheets(children, (s) => (layouts.has(s.id) ? { ...s, layout: layouts.get(s.id)! } : s)),
          );
          setPathSheetId(null);
          toast.success(`Updated the drawing path of ${plural(layouts.size, "sheet")}`);
        }}
      />

      <MoveDialog
        ids={moveIds}
        tree={model.children}
        onOpenChange={(open) => !open && setMoveIds(null)}
        onMove={(ids, folderId) => {
          move(ids, folderId, "inside");
          setMoveIds(null);
        }}
      />

      <AlertDialog open={deleteIds !== null} onOpenChange={(open) => !open && setDeleteIds(null)}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>
              Remove{" "}
              {[
                deleteSummary.sheets && plural(deleteSummary.sheets, "sheet"),
                deleteSummary.folders && plural(deleteSummary.folders, "folder"),
              ]
                .filter(Boolean)
                .join(" and ")}
              ?
            </AlertDialogTitle>
            <AlertDialogDescription asChild>
              <div className="space-y-2">
                {deleteSummary.sheets > 0 && (
                  <p>
                    The sheets are taken out of this sheet set. Their drawings and layouts aren't deleted, but AutoCAD
                    will no longer list them here.
                  </p>
                )}
                {deleteSummary.folders > 0 && (
                  <p>Everything inside the removed folders is kept and moves up one level.</p>
                )}
                <p>You can undo this until you close the page.</p>
              </div>
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancel</AlertDialogCancel>
            <AlertDialogAction onClick={confirmDelete} className="bg-destructive text-white hover:bg-destructive/90">
              Remove
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
