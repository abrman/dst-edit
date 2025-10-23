import type { SheetData } from "./types";

const EXCLUDED_FIELDS = ["ID", "Flags", "Value", "AcDbHandle", "Environ_FileName", "Relative_FileName", "Name"];

const COLUMN_ORDER = ["Number", "Title", "Desc", "FileName"];

export function parseDSTFile(xmlString: string): SheetData[] {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlString, "text/xml");

  const sheets: SheetData[] = [];
  const sheetElements = xmlDoc.getElementsByTagName("AcSmSheet");

  for (let i = 0; i < sheetElements.length; i++) {
    const sheetElement = sheetElements[i];
    const sheetData: SheetData = {};

    // Get sheet ID
    const id = sheetElement.getAttribute("ID");
    if (id) sheetData["ID"] = id;

    // Get all AcSmProp elements within this sheet
    const props = sheetElement.getElementsByTagName("AcSmProp");
    for (let j = 0; j < props.length; j++) {
      const prop = props[j];
      const propName = prop.getAttribute("propname");
      const propValue = prop.textContent || "";

      if (propName) {
        if (propName === "Name") {
          sheetData["Name"] = propValue;
        } else if (!EXCLUDED_FIELDS.includes(propName)) {
          sheetData[propName] = propValue;
        }
      }
    }

    // Get custom properties
    const customBags = sheetElement.getElementsByTagName("AcSmCustomPropertyBag");
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
              sheetData[propName] = innerProp.textContent || "";
            }
          }
        }
      }
    }

    sheets.push(sheetData);
  }

  return sheets;
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

export function reconstructXML(originalXML: string, sheets: SheetData[]): string {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(originalXML, "text/xml");

  const sheetElements = xmlDoc.getElementsByTagName("AcSmSheet");

  for (let i = 0; i < sheetElements.length && i < sheets.length; i++) {
    const sheetElement = sheetElements[i];
    const sheetData = sheets[i];

    const generatedName = `${sheetData["Number"] || ""} ${sheetData["Title"] || ""}`.trim();

    // Update AcSmProp elements
    const props = sheetElement.getElementsByTagName("AcSmProp");
    for (let j = 0; j < props.length; j++) {
      const prop = props[j];
      const propName = prop.getAttribute("propname");

      if (propName) {
        if (propName === "Name") {
          prop.textContent = generatedName;
        } else if (propName in sheetData) {
          prop.textContent = sheetData[propName];
        }
      }
    }

    // Update custom properties
    const customBags = sheetElement.getElementsByTagName("AcSmCustomPropertyBag");
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
            }
          }
        }
      }
    }
  }

  const serializer = new XMLSerializer();
  return serializer.serializeToString(xmlDoc);
}
