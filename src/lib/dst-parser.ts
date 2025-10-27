// lib/dst-parser.ts
import type { SheetData, PropertyMetadata, ParseResult } from "./types";

const EXCLUDED_FIELDS = ["ID", "Flags", "Value", "AcDbHandle", "Environ_FileName", "Relative_FileName", "Name"];

const COLUMN_ORDER = ["Number", "Title", "Desc", "FileName"];

function generateGUID(): string {
  return "gXXXXXXXX-XXXX-4XXX-YXXX-XXXXXXXXXXXX".replace(/[XY]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "X" ? r : (r & 0x3) | 0x8;
    return v.toString(16).toUpperCase();
  });
}

export function parseDSTFile(xmlString: string): ParseResult {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlString, "text/xml");

  const sheets: SheetData[] = [];
  const propertyMetadata = new Map<string, PropertyMetadata>();
  const sheetElements = xmlDoc.getElementsByTagName("AcSmSheet");

  for (let i = 0; i < sheetElements.length; i++) {
    const sheetElement = sheetElements[i];
    const sheetData: SheetData = {};

    // Get sheet ID
    const id = sheetElement.getAttribute("ID");
    if (id) sheetData["ID"] = id;

    // Get all AcSmProp elements (including nested ones)
    const allProps = sheetElement.getElementsByTagName("AcSmProp");

    // Track which props are inside custom property bag
    const customBags = sheetElement.getElementsByTagName("AcSmCustomPropertyBag");
    const customPropElements = new Set<Element>();

    if (customBags.length > 0) {
      const customProps = customBags[0].getElementsByTagName("AcSmProp");
      for (let k = 0; k < customProps.length; k++) {
        customPropElements.add(customProps[k]);
      }
    }

    // Process regular properties (not in custom bag)
    for (let j = 0; j < allProps.length; j++) {
      const prop = allProps[j];

      // Skip if this is inside a custom property bag
      if (customPropElements.has(prop)) continue;

      const propName = prop.getAttribute("propname");
      const propValue = prop.textContent || "";
      const vt = prop.getAttribute("vt") || "8";

      if (propName) {
        if (propName === "Name") {
          sheetData["Name"] = propValue;
        } else if (!EXCLUDED_FIELDS.includes(propName)) {
          sheetData[propName] = propValue;
          if (!propertyMetadata.has(propName)) {
            propertyMetadata.set(propName, { isCustom: false, vt });
          }
        }
      }
    }

    // Get custom properties
    if (customBags.length > 0) {
      const customProps = customBags[0].getElementsByTagName("AcSmCustomPropertyValue");
      for (let k = 0; k < customProps.length; k++) {
        const customProp = customProps[k];
        const propName = customProp.getAttribute("propname");

        if (propName && !EXCLUDED_FIELDS.includes(propName)) {
          const innerProps = customProp.getElementsByTagName("AcSmProp");
          for (let l = 0; l < innerProps.length; l++) {
            const innerProp = innerProps[l];
            const innerPropName = innerProp.getAttribute("propname");
            if (innerPropName === "Value") {
              const vt = innerProp.getAttribute("vt") || "8";
              sheetData[propName] = innerProp.textContent || "";
              if (!propertyMetadata.has(propName)) {
                propertyMetadata.set(propName, { isCustom: true, vt });
              }
            }
          }
        }
      }
    }

    sheets.push(sheetData);
  }

  return { sheets, propertyMetadata };
}

export function reorderColumns(sheets: SheetData[]): SheetData[] {
  if (sheets.length === 0) return sheets;

  return sheets.map((sheet) => {
    const reordered: SheetData = {};

    // Add columns in preferred order first
    COLUMN_ORDER.forEach((col) => {
      if (col in sheet) {
        reordered[col] = sheet[col];
      }
    });

    // Add remaining columns
    Object.keys(sheet).forEach((key) => {
      if (!COLUMN_ORDER.includes(key)) {
        reordered[key] = sheet[key];
      }
    });

    return reordered;
  });
}

export function reconstructXML(
  originalXML: string,
  sheets: SheetData[],
  propertyMetadata: Map<string, PropertyMetadata>
): string {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(originalXML, "text/xml");

  const sheetElements = xmlDoc.getElementsByTagName("AcSmSheet");

  for (let i = 0; i < sheetElements.length && i < sheets.length; i++) {
    const sheetElement = sheetElements[i];
    const sheetData = sheets[i];

    const generatedName = `${sheetData["Number"] || ""} ${sheetData["Title"] || ""}`.trim();

    // Get all AcSmProp elements and track which are in custom bags
    const customBags = sheetElement.getElementsByTagName("AcSmCustomPropertyBag");
    const customPropElements = new Set<Element>();

    if (customBags.length > 0) {
      const customProps = customBags[0].getElementsByTagName("AcSmProp");
      for (let k = 0; k < customProps.length; k++) {
        customPropElements.add(customProps[k]);
      }
    }

    // Update existing regular properties (including nested ones like FileName)
    const allProps = sheetElement.getElementsByTagName("AcSmProp");
    const updatedRegularProps = new Set<string>();

    for (let j = 0; j < allProps.length; j++) {
      const prop = allProps[j];

      // Skip if in custom bag
      if (customPropElements.has(prop)) continue;

      const propName = prop.getAttribute("propname");

      if (propName) {
        if (propName === "Name") {
          prop.textContent = generatedName;
        } else if (propName in sheetData) {
          prop.textContent = sheetData[propName];
          updatedRegularProps.add(propName);
        }
      }
    }

    // Update existing custom properties
    const updatedCustomProps = new Set<string>();

    if (customBags.length > 0) {
      const customProps = customBags[0].getElementsByTagName("AcSmCustomPropertyValue");
      for (let k = 0; k < customProps.length; k++) {
        const customProp = customProps[k];
        const propName = customProp.getAttribute("propname");

        if (propName && propName in sheetData) {
          const innerProps = customProp.getElementsByTagName("AcSmProp");
          for (let l = 0; l < innerProps.length; l++) {
            const innerProp = innerProps[l];
            const innerPropName = innerProp.getAttribute("propname");
            if (innerPropName === "Value") {
              innerProp.textContent = sheetData[propName];
              updatedCustomProps.add(propName);
            }
          }
        }
      }
    }

    // Add missing properties
    for (const [propName, value] of Object.entries(sheetData)) {
      if (EXCLUDED_FIELDS.includes(propName) || propName === "ID" || propName === "Name") {
        continue;
      }

      const metadata = propertyMetadata.get(propName);
      const isCustom = metadata?.isCustom ?? false;
      const vt = metadata?.vt ?? "8";

      if (isCustom && !updatedCustomProps.has(propName) && value.trim() !== "") {
        // Add as custom property
        let customBag = customBags[0];
        if (!customBag) {
          // Create custom property bag if it doesn't exist
          customBag = xmlDoc.createElement("AcSmCustomPropertyBag");
          customBag.setAttribute("ID", generateGUID());
          customBag.setAttribute("clsid", "g4D103908-8C86-4D95-BBF4-68B9A7B00731");
          customBag.setAttribute("propname", "CustomPropertyBag");
          customBag.setAttribute("vt", "13");
          sheetElement.insertBefore(customBag, sheetElement.firstChild);
        }

        const customPropValue = xmlDoc.createElement("AcSmCustomPropertyValue");
        customPropValue.setAttribute("ID", generateGUID());
        customPropValue.setAttribute("clsid", "g8D22A2A4-1777-4D78-84CC-69EF741FE954");
        customPropValue.setAttribute("propname", propName);
        customPropValue.setAttribute("vt", "13");

        const flagsProp = xmlDoc.createElement("AcSmProp");
        flagsProp.setAttribute("propname", "Flags");
        flagsProp.setAttribute("vt", "3");
        flagsProp.textContent = "2";

        const valueProp = xmlDoc.createElement("AcSmProp");
        valueProp.setAttribute("propname", "Value");
        valueProp.setAttribute("vt", vt);
        valueProp.textContent = value;

        customPropValue.appendChild(flagsProp);
        customPropValue.appendChild(valueProp);
        customBag.appendChild(customPropValue);
      } else if (!isCustom && !updatedRegularProps.has(propName) && value.trim() !== "") {
        // Add as regular property (direct child of AcSmSheet)
        const newProp = xmlDoc.createElement("AcSmProp");
        newProp.setAttribute("propname", propName);
        newProp.setAttribute("vt", vt);
        newProp.textContent = value;

        // Insert before SheetViews or at the end
        const sheetViews = sheetElement.getElementsByTagName("AcSmSheetViews")[0];
        if (sheetViews) {
          sheetElement.insertBefore(newProp, sheetViews);
        } else {
          sheetElement.appendChild(newProp);
        }
      }
    }
  }

  const serializer = new XMLSerializer();
  return serializer.serializeToString(xmlDoc);
}
