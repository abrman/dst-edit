import { useEffect, useRef, useState } from "react";
import type React from "react";
import {
  ArrowDown,
  ArrowUp,
  ChevronDown,
  ChevronRight,
  Folder,
  FolderInput,
  FolderOpen,
  FolderPlus,
  FolderSymlink,
  GripVertical,
  MoreHorizontal,
  Pencil,
  Trash2,
  TriangleAlert,
} from "lucide-react";
import { Checkbox } from "@/components/ui/checkbox";
import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuShortcut,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { CommitInput } from "@/components/commit-input";
import { allSheets, type DropPosition, type VisibleRow } from "@/lib/tree";
import type { FolderNode, SheetNode } from "@/lib/dst-model";
import { isUserSpecific, userFolderOf } from "@/lib/paths";
import { cleanSheetValue, isTitleTooLong, replaceSlashes, TITLE_MAX_LENGTH } from "@/lib/autocad-text";
import { cn } from "@/lib/utils";

const INDENT = 20;

export const COLUMN_LABELS: Record<string, string> = { Desc: "Description" };

const COLUMN_WIDTHS: Record<string, string> = {
  Number: "w-32 min-w-32",
  Title: "min-w-72",
  Desc: "min-w-56",
};

interface TreeGridProps {
  rows: VisibleRow[];
  columns: string[];
  customColumns: Set<string>;
  original: Map<string, SheetNode>;
  selection: Set<string>;
  onSelectionChange: (selection: Set<string>) => void;
  collapsed: Set<string>;
  onToggleCollapsed: (id: string) => void;
  dragEnabled: boolean;
  focusFolderId: string | null;
  onFolderFocused: () => void;
  onEditSheet: (id: string, column: string, value: string) => void;
  onPasteGrid: (sheetId: string, column: string, grid: string[][]) => void;
  onEditFolder: (id: string, patch: Partial<Pick<FolderNode, "name" | "desc">>) => void;
  onMove: (ids: string[], targetId: string, position: DropPosition) => void;
  onShift: (ids: string[], direction: -1 | 1) => void;
  onMoveTo: (ids: string[]) => void;
  onDelete: (ids: string[]) => void;
  onNewFolderInside: (folderId: string) => void;
  onModifyDrawingPath: (sheetId: string) => void;
}

/** Tab-separated text copied from Excel, as rows of cells. */
function parseClipboardGrid(text: string): string[][] | null {
  if (!/[\t\n]/.test(text.replace(/\r?\n$/, ""))) return null;
  return text
    .replace(/\r?\n$/, "")
    .split(/\r?\n/)
    .map((line) => line.split("\t"));
}

export function TreeGrid(props: TreeGridProps) {
  const { rows, columns, customColumns, original, selection, onSelectionChange, collapsed, dragEnabled } = props;
  const tableRef = useRef<HTMLTableElement>(null);
  const lastClicked = useRef<number | null>(null);
  const dragIds = useRef<string[]>([]);
  const [drop, setDrop] = useState<{ id: string; position: DropPosition } | null>(null);

  // Put the cursor in a newly created folder's name so it can be typed straight away
  useEffect(() => {
    if (!props.focusFolderId) return;
    const input = tableRef.current?.querySelector<HTMLInputElement>(`[data-folder-name="${props.focusFolderId}"]`);
    if (input) {
      input.focus();
      input.select();
      input.scrollIntoView({ block: "nearest" });
      props.onFolderFocused();
    }
  });

  const idsFor = (id: string) => (selection.has(id) ? [...selection] : [id]);

  const toggleSelect = (index: number, shift: boolean) => {
    const id = rows[index].node.id;
    const next = new Set(selection);
    if (shift && lastClicked.current !== null) {
      const [from, to] = [Math.min(lastClicked.current, index), Math.max(lastClicked.current, index)];
      const select = !selection.has(id);
      for (let i = from; i <= to; i++) {
        if (select) next.add(rows[i].node.id);
        else next.delete(rows[i].node.id);
      }
    } else if (next.has(id)) {
      next.delete(id);
    } else {
      next.add(id);
    }
    lastClicked.current = index;
    onSelectionChange(next);
  };

  const allSelected = rows.length > 0 && rows.every((r) => selection.has(r.node.id));
  const someSelected = rows.some((r) => selection.has(r.node.id));

  /** Moves focus to the same column in the row above or below. */
  const focusVertical = (input: HTMLInputElement, direction: -1 | 1) => {
    const col = input.dataset.col;
    const inputs = Array.from(tableRef.current?.querySelectorAll<HTMLInputElement>(`input[data-col="${col}"]`) ?? []);
    const next = inputs[inputs.indexOf(input) + direction];
    if (next) {
      next.focus();
      next.select();
    }
  };

  const cellKeyDown = (id: string) => (e: React.KeyboardEvent<HTMLInputElement>) => {
    const input = e.currentTarget;
    if (e.altKey && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
      e.preventDefault();
      input.blur();
      props.onShift(idsFor(id), e.key === "ArrowUp" ? -1 : 1);
      return;
    }
    if (e.key === "ArrowUp" || e.key === "ArrowDown" || e.key === "Enter") {
      e.preventDefault();
      input.blur(); // commits the edit
      focusVertical(input, e.key === "ArrowUp" || (e.key === "Enter" && e.shiftKey) ? -1 : 1);
    }
  };

  const onDragOver = (e: React.DragEvent, row: VisibleRow) => {
    if (!dragIds.current.length) return;
    e.preventDefault();
    const rect = e.currentTarget.getBoundingClientRect();
    const y = (e.clientY - rect.top) / rect.height;
    const position: DropPosition =
      row.node.kind === "folder" ? (y < 0.25 ? "before" : y > 0.75 ? "after" : "inside") : y < 0.5 ? "before" : "after";
    if (drop?.id !== row.node.id || drop.position !== position) setDrop({ id: row.node.id, position });
  };

  const endDrag = () => {
    dragIds.current = [];
    setDrop(null);
  };

  const rowMenu = (row: VisibleRow) => {
    const ids = idsFor(row.node.id);
    const many = ids.length > 1 ? ` (${ids.length})` : "";
    const isFolder = row.node.kind === "folder";
    return (
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <Button variant="ghost" size="icon-sm" className="size-7 opacity-60 hover:opacity-100" aria-label="Row actions">
            <MoreHorizontal />
          </Button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="end" className="w-60">
          {isFolder && (
            <>
              <DropdownMenuItem
                onSelect={() =>
                  tableRef.current?.querySelector<HTMLInputElement>(`[data-folder-name="${row.node.id}"]`)?.select()
                }
              >
                <Pencil /> Rename
              </DropdownMenuItem>
              <DropdownMenuItem onSelect={() => props.onNewFolderInside(row.node.id)}>
                <FolderPlus /> New folder inside
              </DropdownMenuItem>
              <DropdownMenuSeparator />
            </>
          )}
          {row.node.kind === "sheet" && isUserSpecific(row.node.layout) && (
            <>
              <DropdownMenuItem onSelect={() => props.onModifyDrawingPath(row.node.id)}>
                <FolderSymlink /> Modify drawing file path…
              </DropdownMenuItem>
              <DropdownMenuSeparator />
            </>
          )}
          <DropdownMenuItem onSelect={() => props.onShift(ids, -1)}>
            <ArrowUp /> Move up{many}
            <DropdownMenuShortcut>Alt+↑</DropdownMenuShortcut>
          </DropdownMenuItem>
          <DropdownMenuItem onSelect={() => props.onShift(ids, 1)}>
            <ArrowDown /> Move down{many}
            <DropdownMenuShortcut>Alt+↓</DropdownMenuShortcut>
          </DropdownMenuItem>
          <DropdownMenuItem onSelect={() => props.onMoveTo(ids)}>
            <FolderInput /> Move to folder…{many}
          </DropdownMenuItem>
          <DropdownMenuSeparator />
          <DropdownMenuItem variant="destructive" onSelect={() => props.onDelete(ids)}>
            <Trash2 />
            {ids.length > 1 ? `Remove selected${many}` : isFolder ? "Remove folder (keep contents)" : "Remove sheet"}
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
    );
  };

  const handleCell = (row: VisibleRow) => (
    <td className="w-6 px-0">
      {dragEnabled && (
        <span
          draggable
          onDragStart={(e) => {
            dragIds.current = idsFor(row.node.id);
            e.dataTransfer.effectAllowed = "move";
            e.dataTransfer.setData("text/plain", row.node.id);
            const tr = e.currentTarget.closest("tr");
            if (tr) e.dataTransfer.setDragImage(tr, 16, 16);
          }}
          onDragEnd={endDrag}
          className="flex h-8 cursor-grab items-center justify-center text-muted-foreground/50 hover:text-foreground active:cursor-grabbing"
          title="Drag to reorder"
        >
          <GripVertical className="size-4" />
        </span>
      )}
    </td>
  );

  return (
    <div className="max-h-[calc(100vh-15rem)] min-h-64 overflow-auto rounded-lg border">
      <table ref={tableRef} className="w-full border-collapse text-sm">
        <thead className="sticky top-0 z-10 bg-muted text-left text-xs font-medium text-muted-foreground shadow-[0_1px_0_0_var(--color-border)]">
          <tr>
            <th className="w-9 px-2.5 py-2">
              <Checkbox
                checked={allSelected ? true : someSelected ? "indeterminate" : false}
                onCheckedChange={() =>
                  onSelectionChange(allSelected ? new Set() : new Set(rows.map((r) => r.node.id)))
                }
                aria-label="Select all"
              />
            </th>
            <th className="w-6" />
            {columns.map((col) => (
              <th key={col} className={cn("px-2 py-2 whitespace-nowrap", COLUMN_WIDTHS[col] ?? "min-w-28")}>
                {COLUMN_LABELS[col] ?? col}
                {customColumns.has(col) && (
                  <span className="ml-1.5 rounded bg-background px-1 py-px text-[10px] font-normal" title="Custom property">
                    custom
                  </span>
                )}
              </th>
            ))}
            <th className="min-w-56 px-2 py-2 whitespace-nowrap">Drawing</th>
            <th className="w-10" />
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => {
            const { node, depth } = row;
            const selected = selection.has(node.id);
            const dropHere = drop?.id === node.id ? drop.position : null;
            const rowClass = cn(
              "group border-b last:border-b-0",
              selected && "bg-primary/6",
              dropHere === "before" && "shadow-[inset_0_2px_0_0_var(--color-primary)]",
              dropHere === "after" && "shadow-[inset_0_-2px_0_0_var(--color-primary)]",
              dropHere === "inside" && "bg-primary/10 outline-2 -outline-offset-2 outline-primary",
            );
            const dragProps = {
              onDragOver: (e: React.DragEvent) => onDragOver(e, row),
              onDragLeave: (e: React.DragEvent) => {
                if (!e.currentTarget.contains(e.relatedTarget as Node)) setDrop(null);
              },
              onDrop: (e: React.DragEvent) => {
                e.preventDefault();
                if (drop && dragIds.current.length) props.onMove(dragIds.current, drop.id, drop.position);
                endDrag();
              },
            };
            const checkbox = (
              <td className="px-2.5">
                <Checkbox
                  checked={selected}
                  onClick={(e) => {
                    e.preventDefault();
                    toggleSelect(index, e.shiftKey);
                  }}
                  aria-label="Select row"
                />
              </td>
            );

            if (node.kind === "folder") {
              const isCollapsed = collapsed.has(node.id);
              const count = allSheets(node.children).length;
              return (
                <tr key={node.id} className={cn(rowClass, !selected && "bg-muted/40")} {...dragProps}>
                  {checkbox}
                  {handleCell(row)}
                  <td colSpan={columns.length + 1} className="py-0.5 pr-2">
                    <div className="flex items-center gap-1" style={{ paddingLeft: depth * INDENT }}>
                      <button
                        type="button"
                        onClick={() => props.onToggleCollapsed(node.id)}
                        className="flex size-6 shrink-0 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-foreground"
                        aria-label={isCollapsed ? "Expand folder" : "Collapse folder"}
                      >
                        {isCollapsed ? <ChevronRight className="size-4" /> : <ChevronDown className="size-4" />}
                      </button>
                      {isCollapsed ? (
                        <Folder className="size-4 shrink-0 text-amber-600" />
                      ) : (
                        <FolderOpen className="size-4 shrink-0 text-amber-600" />
                      )}
                      <CommitInput
                        value={node.name}
                        onCommit={(name) => props.onEditFolder(node.id, { name })}
                        normalize={replaceSlashes}
                        onKeyDown={cellKeyDown(node.id)}
                        data-folder-name={node.id}
                        data-col="__folder"
                        placeholder="Folder name"
                        className="w-80 max-w-[45%] font-semibold"
                      />
                      <CommitInput
                        value={node.desc}
                        onCommit={(desc) => props.onEditFolder(node.id, { desc })}
                        normalize={replaceSlashes}
                        onKeyDown={cellKeyDown(node.id)}
                        data-col="__folderDesc"
                        placeholder="Add description"
                        className="flex-1 text-muted-foreground"
                      />
                      <span className="shrink-0 rounded-full bg-background px-2 py-0.5 text-xs text-muted-foreground tabular-nums">
                        {count} sheet{count === 1 ? "" : "s"}
                      </span>
                    </div>
                  </td>
                  <td className="pr-1 text-right">{rowMenu(row)}</td>
                </tr>
              );
            }

            const before = original.get(node.id);
            const userSpecific = isUserSpecific(node.layout);
            return (
              <tr
                key={node.id}
                className={cn(
                  rowClass,
                  userSpecific ? "bg-orange-50 hover:bg-orange-100/70 dark:bg-orange-500/10" : "hover:bg-muted/30",
                )}
                {...dragProps}
              >
                {checkbox}
                {handleCell(row)}
                {columns.map((col, i) => {
                  const value = node.props[col] ?? "";
                  const changed = (before?.props[col] ?? "") !== value;
                  const tooLong = col === "Title" && isTitleTooLong(value);
                  return (
                    <td
                      key={col}
                      className="py-0.5 pr-1"
                      style={i === 0 ? { paddingLeft: depth * INDENT + (depth ? 30 : 4) } : undefined}
                    >
                      <CommitInput
                        value={value}
                        onCommit={(v) => props.onEditSheet(node.id, col, v)}
                        normalize={(v) => cleanSheetValue(col, v)}
                        onKeyDown={cellKeyDown(node.id)}
                        onPaste={(e) => {
                          const grid = parseClipboardGrid(e.clipboardData.getData("text/plain"));
                          if (!grid) return;
                          e.preventDefault();
                          props.onPasteGrid(node.id, col, grid);
                        }}
                        data-col={col}
                        aria-invalid={tooLong || undefined}
                        className={cn(
                          changed && "bg-amber-100/70 dark:bg-amber-500/15",
                          col === "Number" && "font-medium tabular-nums",
                          tooLong && "border-destructive bg-destructive/10 hover:border-destructive",
                        )}
                        title={
                          tooLong
                            ? `${value.length} characters; AutoCAD allows at most ${TITLE_MAX_LENGTH} in a sheet title`
                            : changed
                              ? `Was: ${before?.props[col] || "(empty)"}`
                              : undefined
                        }
                      />
                      {tooLong && (
                        <span className="block px-2 text-[11px] text-destructive tabular-nums">
                          {value.length}/{TITLE_MAX_LENGTH} characters, too long for AutoCAD
                        </span>
                      )}
                    </td>
                  );
                })}
                <td className="max-w-80 px-2 py-0.5 text-xs">
                  {(() => {
                    const path = node.layout.fileName || node.layout.relative;
                    const drawing = path.split("\\").pop() || "—";
                    const pathChanged = before && before.layout.fileName !== node.layout.fileName;
                    if (userSpecific) {
                      return (
                        <div
                          className="flex items-center gap-1.5 font-medium text-orange-700 dark:text-orange-400"
                          title={`${path}\n\nStored in ${userFolderOf(path) ?? "a user folder"}, so only that user can open it. Use ⋯ › Modify drawing file path to fix.`}
                        >
                          <TriangleAlert className="size-3.5 shrink-0" />
                          <span className="truncate">{drawing}</span>
                        </div>
                      );
                    }
                    return (
                      <div
                        className={cn(
                          "truncate rounded-sm px-1 py-0.5 text-foreground/80",
                          pathChanged && "bg-amber-100/70 dark:bg-amber-500/15",
                        )}
                        title={pathChanged ? `${path}\n\nWas: ${before.layout.fileName}` : path}
                      >
                        {drawing}
                      </div>
                    );
                  })()}
                </td>
                <td className="pr-1 text-right">{rowMenu(row)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
