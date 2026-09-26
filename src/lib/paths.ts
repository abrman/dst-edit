// lib/paths.ts
// Windows path helpers for the sheet set's file references, which store each path three ways:
// absolute (FileName), relative to the .dst (Relative_FileName) and with environment variables
// (Environ_FileName).
import type { SheetLayout, SheetNode } from "./dst-model";

function splitPath(path: string): string[] {
  return path.replace(/\//g, "\\").split("\\").filter((part) => part !== "");
}

const samePart = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

/**
 * The folder the .dst lives in, worked out from sheets whose absolute and relative drawing paths
 * are both known (".\3 - DRAWINGS\G0.dwg" inside "C:\Job\CAD\3 - DRAWINGS\G0.dwg" → "C:\Job\CAD").
 */
export function inferBaseDir(sheets: SheetNode[]): string | null {
  const votes = new Map<string, number>();
  for (const { layout } of sheets) {
    if (!layout.fileName || !layout.relative) continue;
    const rel = splitPath(layout.relative);
    if (rel[0] === ".") rel.shift();
    if (rel.length === 0 || rel.some((part) => part === "..") || /^[A-Za-z]:/.test(rel[0])) continue;
    const abs = splitPath(layout.fileName);
    const tail = abs.slice(abs.length - rel.length);
    if (tail.length !== rel.length || !tail.every((part, i) => samePart(part, rel[i]))) continue;
    const base = abs.slice(0, abs.length - rel.length).join("\\");
    votes.set(base, (votes.get(base) ?? 0) + 1);
  }
  let best: string | null = null;
  for (const [base, count] of votes) if (best === null || count > votes.get(best)!) best = base;
  return best;
}

/** Path of target relative to base, in AutoCAD's style (".\Plots", "..\Plots"); "" on another drive. */
export function relativePath(base: string, target: string): string {
  const from = splitPath(base);
  const to = splitPath(target);
  if (!from.length || !to.length || !samePart(from[0], to[0])) return "";
  let common = 0;
  while (common < from.length && common < to.length && samePart(from[common], to[common])) common++;
  const ups = from.length - common;
  const rest = to.slice(common);
  if (ups === 0) return rest.length ? ".\\" + rest.join("\\") : ".";
  return [...Array(ups).fill(".."), ...rest].join("\\");
}

/** The path written with environment variables, the way AutoCAD stores it. */
export function environPath(target: string): string {
  const path = target.replace(/\//g, "\\");
  const profile = userFolderOf(path);
  if (profile) return "%USERPROFILE%" + path.slice(profile.length);
  if (/^C:/i.test(path)) return "%HOMEDRIVE%" + path.slice(2);
  return path;
}

/** The user's profile folder a path is inside ("C:\Users\Paul"), or null. */
export function userFolderOf(path: string): string | null {
  return /^[A-Za-z]:\\Users\\[^\\]+(?=\\|$)/i.exec(path.replace(/\//g, "\\"))?.[0] ?? null;
}

/** True when a drawing is stored in one user's profile folder, so other users can't open it. */
export function isUserSpecific(layout: SheetLayout): boolean {
  return userFolderOf(layout.fileName) !== null || /^%USERPROFILE%/i.test(layout.environ);
}

/** Replaces a leading folder of path (whole folder names only, any case); null when path isn't inside it. */
export function replacePathPrefix(path: string, find: string, replace: string): string | null {
  const from = splitPath(find);
  const parts = splitPath(path);
  if (!from.length || from.length > parts.length || !from.every((part, i) => samePart(part, parts[i]))) return null;
  return [...splitPath(replace), ...parts.slice(from.length)].join("\\");
}

/**
 * Suggests which part of a user-folder path to swap for a shared one, by finding where the rest of
 * the path appears in the other drawings' paths. For "C:\Users\Paul\Dropbox\Job\CAD\A.dwg" next to
 * drawings in "C:\Projects\Job\CAD\…" it suggests replacing "C:\Users\Paul\Dropbox" with "C:\Projects".
 */
export function suggestSharedPrefix(path: string, others: string[]): { find: string; replace: string } | null {
  const userFolder = userFolderOf(path);
  if (!userFolder) return null;
  const rest = splitPath(path).slice(splitPath(userFolder).length);
  const shared = others.filter((p) => p && !userFolderOf(p)).map(splitPath);

  // Try dropping 0, 1 or 2 folders after the user folder (e.g. "Dropbox"), fewest first
  for (let skip = 0; skip <= 2 && skip < rest.length - 1; skip++) {
    const key = rest.slice(skip, skip + 2);
    const votes = new Map<string, number>();
    for (const parts of shared) {
      const at = parts.findIndex((_, i) => key.every((k, j) => parts[i + j] !== undefined && samePart(parts[i + j], k)));
      if (at > 0) {
        const prefix = parts.slice(0, at).join("\\");
        votes.set(prefix, (votes.get(prefix) ?? 0) + 1);
      }
    }
    const best = [...votes].sort((a, b) => b[1] - a[1])[0];
    if (best) return { find: [userFolder, ...rest.slice(0, skip)].join("\\"), replace: best[0] };
  }
  return null;
}

/**
 * Points a drawing reference at a new folder and updates all the ways AutoCAD stores the path.
 * The relative path is worked out from the sheet set's folder (moved the same way when it was
 * inside the replaced folder too).
 */
export function retargetLayout(layout: SheetLayout, find: string, replace: string, baseDir: string | null): SheetLayout {
  const fileName = replacePathPrefix(layout.fileName, find, replace);
  if (fileName === null) return layout;
  const base = baseDir ? (replacePathPrefix(baseDir, find, replace) ?? baseDir) : null;
  const profile = userFolderOf(fileName);
  return {
    ...layout,
    fileName,
    environ: environPath(fileName),
    relative: (base && relativePath(base, fileName)) || layout.relative,
    specialFolder: profile ? "$(CSIDL_PROFILE)" + fileName.slice(profile.length) : "",
  };
}
