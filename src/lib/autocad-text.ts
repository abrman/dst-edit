// lib/autocad-text.ts
// AutoCAD's limits on sheet text: a sheet title can be at most 64 characters, and "/" isn't
// allowed in titles, descriptions or the layout names built from them.

export const TITLE_MAX_LENGTH = 64;

/** Columns (and folder fields) where "/" is replaced. Number is included because it's part of the layout name. */
const NO_SLASH_COLUMNS = new Set(["Number", "Title", "Desc"]);

// COMBINING LONG SOLIDUS OVERLAY: drawn over the character before it, so after a space it looks like "/"
const SLASH_LOOKALIKE = "̸";

/**
 * Replaces "/" with a look-alike surrounded by single spaces, absorbing any spaces already around it:
 * "Power / Electrical" and "Power/Electrical" both become "Power ̸ Electrical".
 */
export function replaceSlashes(text: string): string {
  return text.replace(/[ \t]*\/[ \t]*/g, ` ${SLASH_LOOKALIKE} `);
}

/** Cleans a value for a sheet column; other columns are returned unchanged. */
export function cleanSheetValue(column: string, value: string): string {
  return NO_SLASH_COLUMNS.has(column) ? replaceSlashes(value) : value;
}

export function isTitleTooLong(title: string | undefined): boolean {
  return (title ?? "").length > TITLE_MAX_LENGTH;
}
