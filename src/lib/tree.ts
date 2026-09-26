// lib/tree.ts
// Immutable operations on the sheet/folder tree.
import type { FolderNode, SheetNode, TreeNode } from "./dst-model";

export type DropPosition = "before" | "after" | "inside";

export interface VisibleRow {
  node: TreeNode;
  depth: number;
}

export function allSheets(nodes: TreeNode[]): SheetNode[] {
  return nodes.flatMap((n) => (n.kind === "sheet" ? [n] : allSheets(n.children)));
}

export function allFolders(nodes: TreeNode[]): FolderNode[] {
  return nodes.flatMap((n) => (n.kind === "folder" ? [n, ...allFolders(n.children)] : []));
}

export function findNode(nodes: TreeNode[], id: string): TreeNode | undefined {
  for (const node of nodes) {
    if (node.id === id) return node;
    if (node.kind === "folder") {
      const found = findNode(node.children, id);
      if (found) return found;
    }
  }
  return undefined;
}

/** Ids of the node and everything inside it. */
export function subtreeIds(node: TreeNode): string[] {
  return node.kind === "sheet" ? [node.id] : [node.id, ...node.children.flatMap(subtreeIds)];
}

/** Ids in tree order, so a multi-row operation keeps rows in the order they're shown. */
export function orderedIds(nodes: TreeNode[], ids: Set<string>): string[] {
  return nodes.flatMap((n) => [
    ...(ids.has(n.id) ? [n.id] : []),
    ...(n.kind === "folder" ? orderedIds(n.children, ids) : []),
  ]);
}

/** Drops ids that are inside another selected folder, since they move with it. */
function topLevelIds(nodes: TreeNode[], ids: Set<string>): string[] {
  const result: string[] = [];
  const walk = (list: TreeNode[]) => {
    for (const n of list) {
      if (ids.has(n.id)) result.push(n.id);
      else if (n.kind === "folder") walk(n.children);
    }
  };
  walk(nodes);
  return result;
}

export function mapSheets(nodes: TreeNode[], fn: (sheet: SheetNode) => SheetNode): TreeNode[] {
  return nodes.map((n) => (n.kind === "sheet" ? fn(n) : { ...n, children: mapSheets(n.children, fn) }));
}

export function updateNode(nodes: TreeNode[], id: string, fn: (node: TreeNode) => TreeNode): TreeNode[] {
  return nodes.map((n) => {
    if (n.id === id) return fn(n);
    return n.kind === "folder" ? { ...n, children: updateNode(n.children, id, fn) } : n;
  });
}

function removeIds(nodes: TreeNode[], ids: Set<string>): TreeNode[] {
  return nodes
    .filter((n) => !ids.has(n.id))
    .map((n) => (n.kind === "folder" ? { ...n, children: removeIds(n.children, ids) } : n));
}

function insertAt(nodes: TreeNode[], targetId: string | null, position: DropPosition, items: TreeNode[]): TreeNode[] {
  if (targetId === null) return [...nodes, ...items];
  const result: TreeNode[] = [];
  for (const n of nodes) {
    if (n.id === targetId) {
      if (position === "before") result.push(...items, n);
      else if (position === "after") result.push(n, ...items);
      else if (n.kind === "folder") result.push({ ...n, children: [...n.children, ...items] });
      else result.push(n, ...items);
    } else {
      result.push(n.kind === "folder" ? { ...n, children: insertAt(n.children, targetId, position, items) } : n);
    }
  }
  return result;
}

/**
 * Moves nodes next to (or into) a target. A null target means the end of the top level.
 * Returns the tree unchanged when the move is impossible (e.g. a folder into itself).
 */
export function moveNodes(
  nodes: TreeNode[],
  ids: string[],
  targetId: string | null,
  position: DropPosition,
): TreeNode[] {
  const moving = topLevelIds(nodes, new Set(ids));
  if (moving.length === 0) return nodes;
  const blocked = new Set(moving.flatMap((id) => subtreeIds(findNode(nodes, id)!)));
  if (targetId !== null && blocked.has(targetId)) return nodes;
  const items = moving.map((id) => findNode(nodes, id)!);
  return insertAt(removeIds(nodes, new Set(moving)), targetId, position, items);
}

/** Moves each node one step up or down among its siblings. */
export function shiftNodes(nodes: TreeNode[], ids: Set<string>, direction: -1 | 1): TreeNode[] {
  const list = [...nodes];
  const indexes = list.map((n, i) => (ids.has(n.id) ? i : -1)).filter((i) => i >= 0);
  if (direction === 1) indexes.reverse();
  for (const i of indexes) {
    const j = i + direction;
    if (j < 0 || j >= list.length || ids.has(list[j].id)) continue;
    [list[i], list[j]] = [list[j], list[i]];
  }
  return list.map((n) => (n.kind === "folder" ? { ...n, children: shiftNodes(n.children, ids, direction) } : n));
}

/** Deletes the given sheets, and removes the given folders while keeping their contents in place. */
export function deleteNodes(nodes: TreeNode[], ids: Set<string>): TreeNode[] {
  return nodes.flatMap((n): TreeNode[] => {
    if (n.kind === "sheet") return ids.has(n.id) ? [] : [n];
    const children = deleteNodes(n.children, ids);
    return ids.has(n.id) ? children : [{ ...n, children }];
  });
}

/** Rows to render, skipping the contents of collapsed folders and rows that don't match the filter. */
export function visibleRows(nodes: TreeNode[], collapsed: Set<string>, filter: string): VisibleRow[] {
  const query = filter.trim().toLowerCase();
  const matches = (sheet: SheetNode) =>
    [...Object.values(sheet.props), sheet.layout.name, sheet.layout.fileName].some((v) =>
      v.toLowerCase().includes(query),
    );

  const walk = (list: TreeNode[], depth: number): VisibleRow[] =>
    list.flatMap((node): VisibleRow[] => {
      if (node.kind === "sheet") return !query || matches(node) ? [{ node, depth }] : [];
      if (query) {
        // A matching folder shows everything in it; otherwise it shows only when something inside matches
        if (node.name.toLowerCase().includes(query)) return [{ node, depth }, ...walkAll(node.children, depth + 1)];
        const inner = walk(node.children, depth + 1);
        return inner.length ? [{ node, depth }, ...inner] : [];
      }
      return [{ node, depth }, ...(collapsed.has(node.id) ? [] : walk(node.children, depth + 1))];
    });

  const walkAll = (list: TreeNode[], depth: number): VisibleRow[] =>
    list.flatMap((node) => [{ node, depth }, ...(node.kind === "folder" ? walkAll(node.children, depth + 1) : [])]);

  return walk(nodes, 0);
}

/** The sheets an edit applies to: selected sheets plus the sheets inside selected folders, in tree order. */
export function targetSheets(nodes: TreeNode[], selection: Set<string>): SheetNode[] {
  if (selection.size === 0) return allSheets(nodes);
  const ids = new Set<string>();
  for (const id of orderedIds(nodes, selection)) {
    for (const inner of subtreeIds(findNode(nodes, id)!)) ids.add(inner);
  }
  return allSheets(nodes).filter((s) => ids.has(s.id));
}

/** The folder containing a node, or null at the top level. */
export function parentOf(nodes: TreeNode[], id: string, parent: FolderNode | null = null): FolderNode | null | undefined {
  for (const n of nodes) {
    if (n.id === id) return parent;
    if (n.kind === "folder") {
      const found = parentOf(n.children, id, n);
      if (found !== undefined) return found;
    }
  }
  return undefined;
}
