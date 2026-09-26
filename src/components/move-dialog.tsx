import { useEffect, useMemo, useState } from "react";
import type React from "react";
import { Folder, Layers } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import type { TreeNode } from "@/lib/dst-model";
import { findNode, subtreeIds } from "@/lib/tree";
import { cn } from "@/lib/utils";

const TOP_LEVEL = "__top";

interface MoveDialogProps {
  ids: string[] | null;
  tree: TreeNode[];
  onOpenChange: (open: boolean) => void;
  onMove: (ids: string[], folderId: string | null) => void;
}

interface FolderOption {
  id: string;
  name: string;
  depth: number;
}

function folderOptions(nodes: TreeNode[], depth = 0): FolderOption[] {
  return nodes.flatMap((n) =>
    n.kind === "folder" ? [{ id: n.id, name: n.name, depth }, ...folderOptions(n.children, depth + 1)] : [],
  );
}

export function MoveDialog({ ids, tree, onOpenChange, onMove }: MoveDialogProps) {
  const [target, setTarget] = useState<string | null>(null);
  useEffect(() => setTarget(null), [ids]);

  const options = useMemo(() => folderOptions(tree), [tree]);
  // A folder can't move into itself or one of its own subfolders
  const blocked = useMemo(
    () => new Set((ids ?? []).flatMap((id) => { const n = findNode(tree, id); return n ? subtreeIds(n) : []; })),
    [ids, tree],
  );

  const option = (id: string, label: string, depth: number, icon: React.ReactNode) => (
    <button
      key={id}
      type="button"
      disabled={blocked.has(id)}
      onClick={() => setTarget(id)}
      onDoubleClick={() => ids && !blocked.has(id) && onMove(ids, id === TOP_LEVEL ? null : id)}
      className={cn(
        "flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left text-sm hover:bg-accent disabled:pointer-events-none disabled:opacity-40",
        target === id && "bg-primary text-primary-foreground hover:bg-primary",
      )}
      style={{ paddingLeft: 8 + depth * 18 }}
    >
      {icon}
      <span className="truncate">{label || "(unnamed folder)"}</span>
    </button>
  );

  return (
    <Dialog open={ids !== null} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle>Move to folder</DialogTitle>
          <DialogDescription>
            Moves {ids?.length === 1 ? "1 item" : `${ids?.length ?? 0} items`} to the end of the folder you pick.
          </DialogDescription>
        </DialogHeader>
        <div className="max-h-80 overflow-y-auto rounded-md border p-1">
          {option(TOP_LEVEL, "Top level (no folder)", 0, <Layers className="size-4 shrink-0" />)}
          {options.map((o) => option(o.id, o.name, o.depth + 1, <Folder className="size-4 shrink-0" />))}
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)}>
            Cancel
          </Button>
          <Button
            disabled={!target || !ids}
            onClick={() => ids && target && onMove(ids, target === TOP_LEVEL ? null : target)}
          >
            Move
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
