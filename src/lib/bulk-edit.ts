// lib/bulk-edit.ts
// Spreadsheet-style edits applied to one column across many sheets.
import type { SheetNode, TreeNode } from "./dst-model";
import { parentOf } from "./tree";
import { cleanSheetValue } from "./autocad-text";

export type BulkOperation =
  | { kind: "set"; value: string }
  | { kind: "replace"; find: string; replace: string; matchCase: boolean }
  | { kind: "series"; start: string; step: number; perFolder: boolean }
  | { kind: "affix"; prefix: string; suffix: string };

/**
 * Splits a series start around its last number, so "C1.01" counts as C1.02, C1.03…
 * and "A-009" as A-010, keeping the zero padding.
 */
function parseSeriesStart(start: string) {
  const match = /^(.*?)(\d+)(\D*)$/.exec(start);
  if (!match) return { prefix: start, number: 1, width: 1, suffix: "" };
  return { prefix: match[1], number: parseInt(match[2], 10), width: match[2].length, suffix: match[3] };
}

export function seriesValue(start: string, step: number, index: number): string {
  const { prefix, number, width, suffix } = parseSeriesStart(start);
  const n = number + step * index;
  const digits = String(Math.abs(n)).padStart(width, "0");
  return `${prefix}${n < 0 ? "-" : ""}${digits}${suffix}`;
}

function escapeRegExp(text: string) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Returns the new value for each target sheet whose value would change. */
export function computeBulkEdit(
  tree: TreeNode[],
  targets: SheetNode[],
  column: string,
  op: BulkOperation,
): Map<string, string> {
  const result = new Map<string, string>();
  const counters = new Map<string, number>();

  targets.forEach((sheet, index) => {
    const current = sheet.props[column] ?? "";
    let next = current;
    switch (op.kind) {
      case "set":
        next = op.value;
        break;
      case "replace":
        if (op.find) {
          next = current.replace(new RegExp(escapeRegExp(op.find), op.matchCase ? "g" : "gi"), () => op.replace);
        }
        break;
      case "series": {
        let position = index;
        if (op.perFolder) {
          const key = parentOf(tree, sheet.id)?.id ?? "";
          position = counters.get(key) ?? 0;
          counters.set(key, position + 1);
        }
        next = seriesValue(op.start, op.step, position);
        break;
      }
      case "affix":
        next = op.prefix + current + op.suffix;
        break;
    }
    next = cleanSheetValue(column, next);
    if (next !== current) result.set(sheet.id, next);
  });

  return result;
}
