import { useEffect, useMemo, useState } from "react";
import { ArrowRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { computeBulkEdit, type BulkOperation } from "@/lib/bulk-edit";
import type { SheetNode, TreeNode } from "@/lib/dst-model";
import { COLUMN_LABELS } from "@/components/tree-grid";
import { isTitleTooLong, TITLE_MAX_LENGTH } from "@/lib/autocad-text";

type Mode = BulkOperation["kind"];

interface BulkEditDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  tree: TreeNode[];
  targets: SheetNode[];
  columns: string[];
  initialColumn: string;
  scopeLabel: string;
  onApply: (column: string, changes: Map<string, string>) => void;
}

const PREVIEW_ROWS = 8;

export function BulkEditDialog(props: BulkEditDialogProps) {
  const { open, targets, columns, tree } = props;
  const [column, setColumn] = useState(props.initialColumn);
  const [mode, setMode] = useState<Mode>("series");
  const [value, setValue] = useState("");
  const [find, setFind] = useState("");
  const [replace, setReplace] = useState("");
  const [matchCase, setMatchCase] = useState(false);
  const [start, setStart] = useState("1");
  const [step, setStep] = useState("1");
  const [perFolder, setPerFolder] = useState(false);
  const [prefix, setPrefix] = useState("");
  const [suffix, setSuffix] = useState("");

  useEffect(() => {
    if (open) setColumn(props.initialColumn);
  }, [open, props.initialColumn]);

  // Start the series from the first target's current value, so "continue numbering" is the default
  useEffect(() => {
    if (open) setStart(targets[0]?.props[column] || "1");
  }, [open, column, targets]);

  const operation: BulkOperation = useMemo(() => {
    switch (mode) {
      case "set":
        return { kind: "set", value };
      case "replace":
        return { kind: "replace", find, replace, matchCase };
      case "affix":
        return { kind: "affix", prefix, suffix };
      case "series":
        return { kind: "series", start, step: Number(step) || 0, perFolder };
    }
  }, [mode, value, find, replace, matchCase, start, step, perFolder, prefix, suffix]);

  const changes = useMemo(
    () => computeBulkEdit(tree, targets, column, operation),
    [tree, targets, column, operation],
  );
  const preview = targets.filter((s) => changes.has(s.id)).slice(0, PREVIEW_ROWS);
  const tooLong = column === "Title" ? [...changes.values()].filter(isTitleTooLong).length : 0;
  const label = (col: string) => COLUMN_LABELS[col] ?? col;

  return (
    <Dialog open={open} onOpenChange={props.onOpenChange}>
      <DialogContent className="sm:max-w-2xl">
        <DialogHeader>
          <DialogTitle>Edit a column</DialogTitle>
          <DialogDescription>Applies to {props.scopeLabel}, in the order they're listed.</DialogDescription>
        </DialogHeader>

        <div className="grid gap-4">
          <div className="grid grid-cols-[8rem_1fr] items-center gap-3">
            <Label>Column</Label>
            <Select value={column} onValueChange={setColumn}>
              <SelectTrigger className="w-full">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                {columns.map((c) => (
                  <SelectItem key={c} value={c}>
                    {label(c)}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <Tabs value={mode} onValueChange={(m) => setMode(m as Mode)}>
            <TabsList className="w-full">
              <TabsTrigger value="series">Number series</TabsTrigger>
              <TabsTrigger value="set">Set value</TabsTrigger>
              <TabsTrigger value="replace">Find &amp; replace</TabsTrigger>
              <TabsTrigger value="affix">Prefix / suffix</TabsTrigger>
            </TabsList>
          </Tabs>

          {mode === "series" && (
            <div className="grid gap-3">
              <div className="grid grid-cols-[8rem_1fr] items-center gap-3">
                <Label htmlFor="series-start">Start at</Label>
                <Input id="series-start" value={start} onChange={(e) => setStart(e.target.value)} />
                <Label htmlFor="series-step">Step</Label>
                <Input
                  id="series-step"
                  type="number"
                  value={step}
                  onChange={(e) => setStep(e.target.value)}
                  className="w-28"
                />
                <Label htmlFor="series-folder">Restart per folder</Label>
                <div className="flex items-center gap-2">
                  <Switch id="series-folder" checked={perFolder} onCheckedChange={setPerFolder} />
                  <span className="text-xs text-muted-foreground">Each folder starts again from the start value</span>
                </div>
              </div>
              <p className="text-xs text-muted-foreground">
                The last number in the start value counts up and keeps its zero padding: <code>1</code> → 2, 3…,{" "}
                <code>C1.01</code> → C1.02, C1.03…, <code>A-009</code> → A-010.
              </p>
            </div>
          )}

          {mode === "set" && (
            <div className="grid grid-cols-[8rem_1fr] items-center gap-3">
              <Label htmlFor="set-value">New value</Label>
              <Input id="set-value" value={value} onChange={(e) => setValue(e.target.value)} placeholder="Leave empty to clear" />
            </div>
          )}

          {mode === "replace" && (
            <div className="grid grid-cols-[8rem_1fr] items-center gap-3">
              <Label htmlFor="find">Find</Label>
              <Input id="find" value={find} onChange={(e) => setFind(e.target.value)} />
              <Label htmlFor="replace">Replace with</Label>
              <Input id="replace" value={replace} onChange={(e) => setReplace(e.target.value)} />
              <Label htmlFor="match-case">Match case</Label>
              <Switch id="match-case" checked={matchCase} onCheckedChange={setMatchCase} />
            </div>
          )}

          {mode === "affix" && (
            <div className="grid grid-cols-[8rem_1fr] items-center gap-3">
              <Label htmlFor="prefix">Add before</Label>
              <Input id="prefix" value={prefix} onChange={(e) => setPrefix(e.target.value)} />
              <Label htmlFor="suffix">Add after</Label>
              <Input id="suffix" value={suffix} onChange={(e) => setSuffix(e.target.value)} />
            </div>
          )}

          <div className="rounded-md border">
            <div className="flex items-center justify-between border-b bg-muted/50 px-3 py-2 text-xs font-medium text-muted-foreground">
              <span>Preview</span>
              <span>
                {changes.size} of {targets.length} sheet{targets.length === 1 ? "" : "s"} change
              </span>
            </div>
            {preview.length === 0 ? (
              <p className="px-3 py-4 text-center text-sm text-muted-foreground">Nothing changes yet.</p>
            ) : (
              <ul className="divide-y text-sm">
                {preview.map((sheet) => (
                  <li key={sheet.id} className="grid grid-cols-[1fr_auto_1fr] items-center gap-3 px-3 py-1.5">
                    <span className="truncate text-muted-foreground line-through decoration-muted-foreground/40">
                      {sheet.props[column] || "(empty)"}
                    </span>
                    <ArrowRight className="size-3.5 text-muted-foreground" />
                    <span className="truncate font-medium">{changes.get(sheet.id) || "(empty)"}</span>
                  </li>
                ))}
                {changes.size > PREVIEW_ROWS && (
                  <li className="px-3 py-1.5 text-xs text-muted-foreground">…and {changes.size - PREVIEW_ROWS} more</li>
                )}
              </ul>
            )}
          </div>
          {tooLong > 0 && (
            <p className="text-sm text-destructive">
              {tooLong} title{tooLong === 1 ? "" : "s"} would be longer than AutoCAD's {TITLE_MAX_LENGTH}-character limit.
              You'll need to shorten {tooLong === 1 ? "it" : "them"} before saving.
            </p>
          )}
        </div>

        <DialogFooter>
          <Button variant="outline" onClick={() => props.onOpenChange(false)}>
            Cancel
          </Button>
          <Button
            disabled={changes.size === 0}
            onClick={() => {
              props.onApply(column, changes);
              props.onOpenChange(false);
            }}
          >
            Update {changes.size} sheet{changes.size === 1 ? "" : "s"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
