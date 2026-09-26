import { useEffect, useMemo, useState } from "react";
import { TriangleAlert } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import type { SheetLayout, SheetNode } from "@/lib/dst-model";
import { replacePathPrefix, retargetLayout, suggestSharedPrefix, userFolderOf } from "@/lib/paths";

interface DrawingPathDialogProps {
  sheet: SheetNode | null;
  sheets: SheetNode[];
  baseDir: string | null;
  onOpenChange: (open: boolean) => void;
  onApply: (layouts: Map<string, SheetLayout>) => void;
}

function PathLine({ label, value }: { label: string; value: string }) {
  return (
    <div className="grid gap-0.5">
      <span className="text-xs text-muted-foreground">{label}</span>
      <code className="rounded bg-muted px-2 py-1 text-xs break-all">{value || "—"}</code>
    </div>
  );
}

export function DrawingPathDialog({ sheet, sheets, baseDir, onOpenChange, onApply }: DrawingPathDialogProps) {
  const [find, setFind] = useState("");
  const [replace, setReplace] = useState("");
  const [applyToAll, setApplyToAll] = useState(true);

  const current = sheet?.layout.fileName ?? "";
  const userFolder = userFolderOf(current);

  // Start from a suggestion based on where the other drawings live
  useEffect(() => {
    if (!sheet) return;
    const suggestion = suggestSharedPrefix(
      sheet.layout.fileName,
      sheets.map((s) => s.layout.fileName),
    );
    const parts = sheet.layout.fileName.split("\\");
    const userDepth = (userFolderOf(sheet.layout.fileName) ?? "").split("\\").length;
    setFind(suggestion?.find ?? parts.slice(0, Math.min(userDepth + 1, parts.length - 1)).join("\\"));
    setReplace(suggestion?.replace ?? "");
    setApplyToAll(true);
  }, [sheet, sheets]);

  const findTrimmed = find.trim().replace(/\\+$/, "");
  const replaceTrimmed = replace.trim().replace(/\\+$/, "");
  const matchesCurrent = !!findTrimmed && replacePathPrefix(current, findTrimmed, "X") !== null;
  const validReplace = /^([A-Za-z]:\\|\\\\)/.test(replaceTrimmed + "\\");

  const changes = useMemo(() => {
    const result = new Map<string, SheetLayout>();
    if (!sheet || !matchesCurrent || !validReplace) return result;
    for (const s of applyToAll ? sheets : [sheet]) {
      const layout = retargetLayout(s.layout, findTrimmed, replaceTrimmed, baseDir);
      if (layout !== s.layout && layout.fileName !== s.layout.fileName) result.set(s.id, layout);
    }
    return result;
  }, [sheet, sheets, applyToAll, findTrimmed, replaceTrimmed, baseDir, matchesCurrent, validReplace]);

  const otherMatches = sheet
    ? sheets.filter((s) => s.id !== sheet.id && replacePathPrefix(s.layout.fileName, findTrimmed, "X") !== null).length
    : 0;
  const updated = sheet ? changes.get(sheet.id) : undefined;
  const stillUserSpecific = updated && userFolderOf(updated.fileName);
  // The link each user creates so the shared path points into their own copy of the folder. Only offered
  // when the replaced folder is inside the user folder (not the user folder itself or above it) and the
  // new path is on a local drive, since mklink /J can't create a link on a network share.
  const linkTarget =
    userFolder && replacePathPrefix(findTrimmed, userFolder, "X") !== null && findTrimmed.length > userFolder.length
      ? "%USERPROFILE%" + findTrimmed.slice(userFolder.length)
      : "";
  const localReplace = /^[A-Za-z]:\\[^\\]/.test(replaceTrimmed);

  return (
    <Dialog open={sheet !== null} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-2xl">
        <DialogHeader>
          <DialogTitle>Modify drawing file path</DialogTitle>
          <DialogDescription>
            This drawing is stored in <code>{userFolder}</code>, so only that user can open it from the sheet set.
            Point it at a path every user has, such as a folder link (e.g. <code>C:\Projects</code>) that leads to
            each user's copy of the shared folder.
          </DialogDescription>
        </DialogHeader>

        <div className="grid gap-4">
          <PathLine label="Current path" value={current} />

          <div className="grid gap-3 sm:grid-cols-2">
            <div className="grid gap-1.5">
              <Label htmlFor="path-find">Replace this folder</Label>
              <Input id="path-find" value={find} onChange={(e) => setFind(e.target.value)} aria-invalid={!matchesCurrent} />
              {!matchesCurrent && (
                <p className="text-xs text-destructive">Must be the start of the current path, e.g. {userFolder}\…</p>
              )}
            </div>
            <div className="grid gap-1.5">
              <Label htmlFor="path-replace">With</Label>
              <Input
                id="path-replace"
                value={replace}
                onChange={(e) => setReplace(e.target.value)}
                placeholder="C:\Projects"
                aria-invalid={!!replace && !validReplace}
              />
              {replace && !validReplace && (
                <p className="text-xs text-destructive">Use a full path, like C:\Projects or \\server\share.</p>
              )}
            </div>
          </div>

          <div className="flex items-center gap-2">
            <Switch id="path-all" checked={applyToAll} onCheckedChange={setApplyToAll} />
            <Label htmlFor="path-all" className="font-normal">
              Also update the {otherMatches} other sheet{otherMatches === 1 ? "" : "s"} whose drawing is in this folder
            </Label>
          </div>

          {updated && (
            <div className="grid gap-2 rounded-md border p-3">
              <PathLine label="New path" value={updated.fileName} />
              <PathLine label="Relative to the sheet set" value={updated.relative} />
              <PathLine label="With environment variables" value={updated.environ} />
            </div>
          )}

          {stillUserSpecific && (
            <p className="flex items-center gap-2 text-sm text-orange-700 dark:text-orange-400">
              <TriangleAlert className="size-4 shrink-0" /> The new path is still inside a user folder.
            </p>
          )}

          {linkTarget && validReplace && localReplace && (
            <div className="grid gap-1 text-xs text-muted-foreground">
              <span>
                Each user needs the link on their computer. If it doesn't exist yet, it can be created from a Command
                Prompt with:
              </span>
              <code className="rounded bg-muted px-2 py-1 break-all text-foreground">
                mklink /J "{replaceTrimmed}" "{linkTarget}"
              </code>
              <span>
                The folder that will hold the link must already exist, and <code>{replaceTrimmed}</code> must not. This
                assumes the shared folder has the same name in every user's profile.
              </span>
            </div>
          )}
          {validReplace && !localReplace && replaceTrimmed.startsWith("\\\\") && (
            <p className="text-xs text-muted-foreground">
              Every user needs access to <code>{replaceTrimmed}</code>. Check that the share is reachable from each
              computer before saving.
            </p>
          )}
        </div>

        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)}>
            Cancel
          </Button>
          <Button disabled={changes.size === 0} onClick={() => onApply(changes)}>
            Update {changes.size} sheet{changes.size === 1 ? "" : "s"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
