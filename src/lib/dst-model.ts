// lib/dst-model.ts
// Parses a decoded .dst (sheet set XML) into an editable tree model, and writes
// a model back into the original XML so everything the editor doesn't know
// about (views, callouts, publish settings, IDs) is preserved.

/** A sheet's link to its layout in a drawing. The drawing path is stored four ways, as AutoCAD writes it. */
export interface SheetLayout {
  name: string;
  fileName: string;
  relative: string;
  environ: string;
  specialFolder: string;
}

const LAYOUT_PATH_PROPS: [keyof SheetLayout, string][] = [
  ["environ", "Environ_FileName"],
  ["fileName", "FileName"],
  ["relative", "Relative_FileName"],
  ["specialFolder", "SpecialFolder_FileName"],
];

export interface SheetNode {
  kind: "sheet";
  id: string;
  /** Regular sheet properties (Number, Title, Desc, …) and custom sheet property values. */
  props: Record<string, string>;
  layout: SheetLayout;
}

export interface FolderNode {
  kind: "folder";
  id: string;
  name: string;
  desc: string;
  children: TreeNode[];
}

export type TreeNode = SheetNode | FolderNode;

export interface CustomPropDef {
  name: string;
  /** "sheet": every sheet has its own value (value is the default). "set": one value for the sheet set. */
  scope: "sheet" | "set";
  value: string;
}

export interface OutputDir {
  fileName: string;
  relative: string;
  environ: string;
}

export interface DstModel {
  /** Direct properties of the sheet set: Name (title shown in AutoCAD), Desc, ProjectName, … */
  setProps: Record<string, string>;
  customProps: CustomPropDef[];
  outputDir: OutputDir;
  defaultFilename: string;
  /** Regular (non-custom) sheet property names found in the file, in display order. */
  sheetColumns: string[];
  children: TreeNode[];
}

const CLSID = {
  subset: "g076D548F-B0F5-4FE1-B35D-7F7B73B8D322",
  customBag: "g4D103908-8C86-4D95-BBF4-68B9A7B00731",
  customValue: "g8D22A2A4-1777-4D78-84CC-69EF741FE954",
  publishOptions: "gF57F96E7-0F16-4DC9-8F09-52F7BB389AB6",
  simpleFileRef: "gD15A03C2-C39B-428A-9BBA-C031347C496F",
};

// Custom property Flags: 1 = sheet set property, 2 = sheet property.
const FLAGS_SET = "1";
const FLAGS_SHEET = "2";

const PREFERRED_COLUMNS = ["Number", "Title", "Desc"];
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8"?>';

export function generateID(): string {
  return "g" + crypto.randomUUID().toUpperCase();
}

// ---------------------------------------------------------------------------
// DOM helpers

function childElements(el: Element, tag?: string): Element[] {
  return Array.from(el.children).filter((c) => !tag || c.tagName === tag);
}

function directProp(el: Element, name: string): Element | undefined {
  return childElements(el, "AcSmProp").find((p) => p.getAttribute("propname") === name);
}

function directPropValue(el: Element, name: string): string {
  return directProp(el, name)?.textContent ?? "";
}

function childByPropname(el: Element, tag: string, propname: string): Element | undefined {
  return childElements(el, tag).find((c) => c.getAttribute("propname") === propname);
}

/** AutoCAD writes the children of each element sorted by propname; new ones are inserted the same way. */
function insertSorted(parent: Element, child: Element) {
  const name = (child.getAttribute("propname") ?? "").toLowerCase();
  const before = childElements(parent).find((c) => {
    if (c.tagName === "AcSmSheet" || c.tagName === "AcSmSubset") return true;
    const other = c.getAttribute("propname");
    return other !== null && other.toLowerCase() > name;
  });
  parent.insertBefore(child, before ?? null);
}

function createProp(doc: Document, name: string, value: string, vt = "8"): Element {
  const prop = doc.createElement("AcSmProp");
  prop.setAttribute("propname", name);
  prop.setAttribute("vt", vt);
  prop.textContent = value;
  return prop;
}

/** Sets a direct AcSmProp, creating it (sorted) when missing and the value is non-empty. */
function setDirectProp(el: Element, name: string, value: string) {
  const prop = directProp(el, name);
  if (prop) {
    if (prop.textContent !== value) prop.textContent = value;
  } else if (value !== "") {
    insertSorted(el, createProp(el.ownerDocument, name, value));
  }
}

function createContainer(doc: Document, tag: string, clsid: string, propname: string): Element {
  const el = doc.createElement(tag);
  el.setAttribute("clsid", clsid);
  el.setAttribute("ID", generateID());
  el.setAttribute("propname", propname);
  el.setAttribute("vt", "13");
  return el;
}

function parseXML(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) throw new Error("The file isn't a valid sheet set: " + (error.textContent ?? "XML error"));
  return doc;
}

function sheetSetElement(doc: Document): Element {
  const sheetSet = doc.getElementsByTagName("AcSmSheetSet")[0];
  if (!sheetSet) throw new Error("The file doesn't contain a sheet set (no AcSmSheetSet element).");
  return sheetSet;
}

// ---------------------------------------------------------------------------
// Parsing

function customBag(el: Element): Element | undefined {
  return childByPropname(el, "AcSmCustomPropertyBag", "CustomPropertyBag");
}

function customEntries(bag: Element | undefined) {
  if (!bag) return [];
  return childElements(bag, "AcSmCustomPropertyValue").map((entry) => ({
    el: entry,
    name: entry.getAttribute("propname") ?? "",
    flags: directPropValue(entry, "Flags"),
    value: directPropValue(entry, "Value"),
  }));
}

export function parseDst(xml: string): DstModel {
  const doc = parseXML(xml);
  const sheetSet = sheetSetElement(doc);

  const setProps: Record<string, string> = {};
  for (const prop of childElements(sheetSet, "AcSmProp")) {
    setProps[prop.getAttribute("propname") ?? ""] = prop.textContent ?? "";
  }

  const customProps: CustomPropDef[] = customEntries(customBag(sheetSet)).map((e) => ({
    name: e.name,
    scope: e.flags === FLAGS_SET ? "set" : "sheet",
    value: e.value,
  }));
  const knownCustom = new Set(customProps.map((d) => d.name));
  const columns = new Set<string>();

  const parseSheet = (el: Element): SheetNode => {
    const props: Record<string, string> = {};
    for (const prop of childElements(el, "AcSmProp")) {
      const name = prop.getAttribute("propname") ?? "";
      if (!name || name === "Name" || isBrokenPropname(name)) continue;
      props[name] = prop.textContent ?? "";
      columns.add(name);
    }
    for (const entry of customEntries(customBag(el))) {
      props[entry.name] = entry.value;
      // A sheet property that only exists on sheets, not in the sheet set definition
      if (!knownCustom.has(entry.name)) {
        knownCustom.add(entry.name);
        customProps.push({ name: entry.name, scope: "sheet", value: "" });
      }
    }
    const layoutEl = childByPropname(el, "AcSmAcDbLayoutReference", "Layout");
    return {
      kind: "sheet",
      id: el.getAttribute("ID") ?? generateID(),
      props,
      layout: {
        name: layoutEl ? directPropValue(layoutEl, "Name") : directPropValue(el, "Name"),
        fileName: layoutEl ? directPropValue(layoutEl, "FileName") : "",
        relative: layoutEl ? directPropValue(layoutEl, "Relative_FileName") : "",
        environ: layoutEl ? directPropValue(layoutEl, "Environ_FileName") : "",
        specialFolder: layoutEl ? directPropValue(layoutEl, "SpecialFolder_FileName") : "",
      },
    };
  };

  const parseChildren = (container: Element): TreeNode[] =>
    childElements(container).flatMap((el): TreeNode[] => {
      if (el.tagName === "AcSmSheet") return [parseSheet(el)];
      if (el.tagName === "AcSmSubset") {
        return [
          {
            kind: "folder",
            id: el.getAttribute("ID") ?? generateID(),
            name: directPropValue(el, "Name"),
            desc: directPropValue(el, "Desc"),
            children: parseChildren(el),
          },
        ];
      }
      return [];
    });

  const children = parseChildren(sheetSet);

  const publish = childElements(sheetSet, "AcSmPublishOptions")[0];
  const outputEl = publish ? childByPropname(publish, "AcSmSimpleFileReferece", "DefaultOutputdir") : undefined;

  return {
    setProps,
    customProps,
    outputDir: {
      fileName: outputEl ? directPropValue(outputEl, "FileName") : "",
      relative: outputEl ? directPropValue(outputEl, "Relative_FileName") : "",
      environ: outputEl ? directPropValue(outputEl, "Environ_FileName") : "",
    },
    defaultFilename: publish ? directPropValue(publish, "DefaultFilename") : "",
    sheetColumns: [...PREFERRED_COLUMNS, ...[...columns].filter((c) => !PREFERRED_COLUMNS.includes(c))],
    children,
  };
}

// ---------------------------------------------------------------------------
// Writing

/** Writes custom property entries into owner's bag and removes entries whose name isn't in `allowed`. */
function applyCustomBag(
  owner: Element,
  entries: { name: string; flags: string; value: string }[],
  allowed: Set<string>,
) {
  const doc = owner.ownerDocument;
  let bag = customBag(owner);
  if (!bag) {
    if (entries.length === 0) return;
    bag = createContainer(doc, "AcSmCustomPropertyBag", CLSID.customBag, "CustomPropertyBag");
    insertSorted(owner, bag);
  }
  const existing = new Map(customEntries(bag).map((e) => [e.name, e.el]));
  for (const [name, el] of existing) if (!allowed.has(name)) el.remove();
  for (const entry of entries) {
    let el = existing.get(entry.name);
    if (!el) {
      el = createContainer(doc, "AcSmCustomPropertyValue", CLSID.customValue, entry.name);
      el.appendChild(createProp(doc, "Flags", entry.flags, "3"));
      el.appendChild(createProp(doc, "Value", entry.value));
      bag.appendChild(el);
      continue;
    }
    setDirectProp(el, "Flags", entry.flags);
    const valueProp = directProp(el, "Value");
    if (valueProp) {
      if (valueProp.textContent !== entry.value) valueProp.textContent = entry.value;
    } else {
      el.appendChild(createProp(doc, "Value", entry.value));
    }
  }
}

function applyOutputDir(sheetSet: Element, model: DstModel) {
  const doc = sheetSet.ownerDocument;
  const { fileName, relative, environ } = model.outputDir;
  let publish = childElements(sheetSet, "AcSmPublishOptions")[0];
  const hasValues = fileName || relative || environ || model.defaultFilename;
  if (!publish) {
    if (!hasValues) return;
    publish = createContainer(doc, "AcSmPublishOptions", CLSID.publishOptions, "PublishOptions");
    insertSorted(sheetSet, publish);
  }
  setDirectProp(publish, "DefaultFilename", model.defaultFilename);

  let outputEl = childByPropname(publish, "AcSmSimpleFileReferece", "DefaultOutputdir");
  if (!outputEl) {
    if (!fileName && !relative && !environ) return;
    outputEl = createContainer(doc, "AcSmSimpleFileReferece", CLSID.simpleFileRef, "DefaultOutputdir");
    insertSorted(publish, outputEl);
  }
  const fields: [string, string][] = [
    ["Environ_FileName", environ],
    ["FileName", fileName],
    ["Relative_FileName", relative],
  ];
  for (const [name, value] of fields) {
    if (value) setDirectProp(outputEl, name, value);
    else directProp(outputEl, name)?.remove();
  }
}

/** Gives an element copied from elsewhere in the file fresh IDs, so no ID is used twice. */
function renewIDs(el: Element) {
  for (const node of [el, ...Array.from(el.getElementsByTagName("*"))]) {
    if (node.hasAttribute("ID")) node.setAttribute("ID", generateID());
  }
}

function createSubset(doc: Document, folder: FolderNode, parent: Element): Element {
  const el = doc.createElement("AcSmSubset");
  el.setAttribute("clsid", CLSID.subset);
  el.setAttribute("ID", folder.id);
  // New folders inherit the template layout and new-sheet location of the nearest parent that has them,
  // like AutoCAD does when you create a subset.
  for (const [tag, propname] of [
    ["AcSmAcDbLayoutReference", "DefDwtLayout"],
    ["AcSmFileReference", "NewSheetLocation"],
  ]) {
    for (let owner: Element | null = parent; owner; owner = owner.parentElement) {
      const source = childByPropname(owner, tag, propname);
      if (source) {
        const copy = source.cloneNode(true) as Element;
        renewIDs(copy);
        el.appendChild(copy);
        break;
      }
      if (owner.tagName === "AcSmSheetSet") break;
    }
  }
  return el;
}

/**
 * The layout name a sheet will have once saved. Renaming or renumbering a sheet renames its layout
 * to "Number Title", as AutoCAD names layouts by default; otherwise the layout keeps its name.
 */
export function layoutNameFor(sheet: SheetNode, original: SheetNode | undefined): string {
  const renamed =
    original && (original.props.Number !== sheet.props.Number || original.props.Title !== sheet.props.Title);
  return renamed ? `${sheet.props.Number ?? ""} ${sheet.props.Title ?? ""}`.trim() : sheet.layout.name;
}

/**
 * Old versions of this editor imported CSVs with Windows line endings wrongly and added duplicate
 * properties with a line break in the name (e.g. "PAGE_INDEX\r"). They're dropped when saving.
 */
function isBrokenPropname(name: string): boolean {
  return /[\r\n]/.test(name);
}

/** Counts the broken duplicate properties that saving will remove. */
export function countBrokenProps(xml: string): number {
  const doc = parseXML(xml);
  return Array.from(doc.getElementsByTagName("AcSmSheet")).reduce(
    (n, sheet) => n + childElements(sheet, "AcSmProp").filter((p) => isBrokenPropname(p.getAttribute("propname") ?? "")).length,
    0,
  );
}

function applySheet(el: Element, sheet: SheetNode, original: SheetNode | undefined, model: DstModel) {
  for (const prop of childElements(el, "AcSmProp")) {
    if (isBrokenPropname(prop.getAttribute("propname") ?? "")) prop.remove();
  }
  const customNames = new Set(model.customProps.map((d) => d.name));
  for (const [name, value] of Object.entries(sheet.props)) {
    if (!customNames.has(name)) setDirectProp(el, name, value);
  }
  const sheetDefs = model.customProps.filter((d) => d.scope === "sheet");
  applyCustomBag(
    el,
    sheetDefs
      .filter((d) => d.name in sheet.props)
      .map((d) => ({ name: d.name, flags: FLAGS_SHEET, value: sheet.props[d.name] })),
    new Set(sheetDefs.map((d) => d.name)),
  );

  const layoutEl = childByPropname(el, "AcSmAcDbLayoutReference", "Layout");
  const name = layoutNameFor(sheet, original);
  if (name !== sheet.layout.name) {
    if (layoutEl) setDirectProp(layoutEl, "Name", name);
    if (directProp(el, "Name")) setDirectProp(el, "Name", name);
  }
  if (layoutEl) {
    for (const [key, propname] of LAYOUT_PATH_PROPS) {
      const value = sheet.layout[key];
      if (value) setDirectProp(layoutEl, propname, value);
      else directProp(layoutEl, propname)?.remove();
    }
  }
}

/** Re-indents the document with tabs, one element per line, like AutoCAD writes it. */
function reindent(el: Element, depth: number) {
  const elements = childElements(el);
  if (elements.length === 0) return;
  for (const node of Array.from(el.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE && !(node.textContent ?? "").trim()) node.remove();
  }
  const doc = el.ownerDocument;
  for (const child of elements) {
    el.insertBefore(doc.createTextNode("\n" + "\t".repeat(depth + 1)), child);
    reindent(child, depth + 1);
  }
  el.appendChild(doc.createTextNode("\n" + "\t".repeat(depth)));
}

function collectSheets(nodes: TreeNode[], into = new Map<string, SheetNode>()): Map<string, SheetNode> {
  for (const node of nodes) {
    if (node.kind === "sheet") into.set(node.id, node);
    else collectSheets(node.children, into);
  }
  return into;
}

/** Writes the model into the original XML and returns the new XML. */
export function applyModel(originalXML: string, model: DstModel): string {
  const doc = parseXML(originalXML);
  const sheetSet = sheetSetElement(doc);
  const originalSheets = collectSheets(parseDst(originalXML).children);

  for (const [name, value] of Object.entries(model.setProps)) setDirectProp(sheetSet, name, value);

  applyCustomBag(
    sheetSet,
    model.customProps.map((d) => ({ name: d.name, flags: d.scope === "set" ? FLAGS_SET : FLAGS_SHEET, value: d.value })),
    new Set(model.customProps.map((d) => d.name)),
  );
  applyOutputDir(sheetSet, model);

  // Detach every sheet and folder, then put back the ones still in the model, in model order.
  // Anything not put back has been deleted.
  const nodeElements = new Map<string, Element>();
  for (const el of [
    ...Array.from(doc.getElementsByTagName("AcSmSheet")),
    ...Array.from(doc.getElementsByTagName("AcSmSubset")),
  ]) {
    const id = el.getAttribute("ID");
    if (id) nodeElements.set(id, el);
  }
  for (const el of nodeElements.values()) el.remove();

  const build = (container: Element, nodes: TreeNode[]) => {
    for (const node of nodes) {
      if (node.kind === "sheet") {
        const el = nodeElements.get(node.id);
        if (!el) continue; // Sheets can't be created here, only kept or removed
        applySheet(el, node, originalSheets.get(node.id), model);
        container.appendChild(el);
      } else {
        const el = nodeElements.get(node.id) ?? createSubset(doc, node, container);
        setDirectProp(el, "Desc", node.desc);
        setDirectProp(el, "Name", node.name);
        container.appendChild(el);
        build(el, node.children);
      }
    }
  };
  build(sheetSet, model.children);

  reindent(doc.documentElement, 0);
  return XML_DECLARATION + "\n" + new XMLSerializer().serializeToString(doc.documentElement);
}
