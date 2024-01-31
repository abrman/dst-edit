import React from "react";
import styled from "styled-components";
import { useDropzone } from "react-dropzone";
import { js2xml, xml2js } from "xml-js";
import Spreadsheet from "react-spreadsheet";
import toast from "react-hot-toast";
import writeXlsxFile, { SheetData } from "write-excel-file";

const Wrapper = styled.div`
  display: flex;
  flex-direction: row;
  align-items: center;
  width: 100vw;
  width: 100dvw;
  height: 100vh;
  height: 100dvh;
  overflow: hidden;
  background-color: #eee;
  max-height: 100vh;

  li {
    margin-bottom: 10px;
  }
`;

const SideWrapper = styled.div`
  width: 500px;
  min-width: 500px;
  height: 100%;
  border-right: 1px solid #ccc;
  background-color: #fff;
  padding: 0 40px;
  display: flex;
  flex-direction: column;
  max-height: 100vh;
  overflow-y: auto;
`;

const SSWrapper = styled.div`
  overflow: auto;
  flex: 1;
  height: 100%;
`;

const Dropzone = styled.div<{ $dragActive: boolean }>`
  box-sizing: border-box;
  padding: 2em 4em;
  border-radius: 0.5em;

  max-width: 80dvw;
  width: 100%;
  text-align: center;
  background: #e9edec;
  outline: 3px solid gray;
  margin-bottom: 2em;
  ${({ $dragActive }) => $dragActive && `outline-color: #285dfb;`}
`;

const Download = styled.button`
  margin: 0 auto;
  outline: 3px solid gray;
  margin-bottom: 20px;
`;

// const buildPathString = (path: (number | string)[]) =>
// path.map((part) => (typeof part === "number" ? `[${part}]` : `["${part}"]`)).join("");
// const findPaths = (
// obj: any,
// propName: string,
// propValue: string,
// upAncestryCount: number = 0,
// path = [] as string[]
// ) => {
// let results: string[] = [];

// if (typeof obj === "object" && obj !== null) {
//   for (let key in obj) {
//     if (obj.hasOwnProperty(key)) {
//       let currentPath = [...path, key];
//       if (typeof obj[key] === "object" && obj[key] !== null && !(key === propName && propValue === "_ANY")) {
//         let nestedResults = findPaths(obj[key], propName, propValue, upAncestryCount, currentPath);
//         results = results.concat(nestedResults);
//       } else if (key === propName && (obj[key] === propValue || propValue === "_ANY")) {
//         for (let i = 0; i < upAncestryCount; i++) currentPath.pop();
//         results.push(buildPathString(currentPath));
//       }
//     }
//   }
// }
// return results;
// };
const findObjects = (obj: any, propName: string, propValue: string, path = [] as string[]) => {
  let results: string[] = [];

  if (typeof obj === "object" && obj !== null) {
    for (let key in obj) {
      if (obj.hasOwnProperty(key)) {
        let currentPath = [...path, key];
        if (typeof obj[key] === "object" && obj[key] !== null && !(key === propName && propValue === "_ANY")) {
          let nestedResults = findObjects(obj[key], propName, propValue, currentPath);
          results = results.concat(nestedResults);
        } else if (key === propName && (obj[key] === propValue || propValue === "_ANY")) {
          results.push(obj[key]);
        }
      }
    }
  }
  return results;
};

// prettier-ignore
const flipMap = {131:10,134:9,172:32,175:33,174:34,169:35,168:36,171:37,170:38,165:39,164:40,167:41,166:42,161:43,160:44,163:45,162:46,221:47,220:48,223:49,222:50,217:51,216:52,219:53,218:54,213:55,212:56,215:57,214:58,209:59,208:60,211:61,210:62,205:63,204:64,207:65,206:66,201:67,200:68,203:69,202:70,197:71,196:72,199:73,198:74,193:75,192:76,195:77,194:78,253:79,252:80,255:81,254:82,249:83,248:84,251:85,250:86,245:87,244:88,247:89,246:90,241:91,240:92,243:93,242:94,237:95,236:96,239:97,238:98,233:99,232:100,235:101,234:102,229:103,228:104,231:105,230:106,225:107,224:108,227:109,226:110,29:111,28:112,31:113,30:114,25:115,24:116,27:117,26:118,21:119,20:120,23:121,22:122};
const reversedMap = Object.keys(flipMap).reduce(
  (map, key) => ({ ...map, [(flipMap as any)[key]]: key }),
  {} as { [key: string]: string }
);
function flipBits(byteValue: number, reverse = false) {
  const map = reverse ? reversedMap : flipMap;
  return (map as any)[byteValue] !== undefined ? (map as any)[byteValue] : byteValue;
}

type XMLCustomProp = {
  _attributes: {
    clsid: string;
    ID: string;
    propname: "CustomPropertyBag";
    vt: string;
  };
  AcSmCustomPropertyValue: {
    _attributes: {
      clsid: string;
      ID: string;
      propname: string; // "PROPERTY NAME HERE"
      vt: string;
    };
    AcSmProp: [
      { _attributes: { propname: "Flags"; vt: "3" }; _text: string },
      { _attributes: { propname: "Value"; vt: "8" }; _text: string } // "PROPERTY VALUE HERE"
    ];
  };
};

type XMLLayoutPropName = "AcDbHandle" | "Environ_FileName" | "FileName" | "Name" | "Relative_FileName";
type XMLLayoutProp = { _attributes: { propname: XMLLayoutPropName; vt: "8" }; _text: string };

type XMLSheetPropName = "Category" | "Desc" | "IssuePurpose" | "Number" | "RevisionDate" | "RevisionNumber" | "Title";
type XMLSheetProp = { _attributes: { propname: XMLSheetPropName; vt: "8" }; _text: string };

type XMLSheet = {
  _attributes: { clsid: string; ID: string };
  AcSmProp: XMLSheetProp | XMLSheetProp[];
  AcSmCustomPropertyBag: XMLCustomProp | XMLCustomProp[];
  AcSmAcDbLayoutReference: {
    _attributes: {
      clsid: string;
      ID: string;
      propname: "Layout";
      vt: "13";
    };
    AcSmProp: XMLLayoutProp | XMLLayoutProp[];
  };
  AcSmSheetViews: {
    _attributes: {
      clsid: string;
      ID: string;
      propname: "SheetViews";
      vt: "13";
    };
  };
};

type CustomProp = {
  root: object;
  name: string;
  key: string;
  value: string;
};

type LayoutProp = {
  root: object;
  name: "AcDbHandle" | "Environ_FileName" | "FileName" | "Name" | "Relative_FileName";
  key: string;
  value: string;
};

type Sheet = {
  root: object;
  Category: string;
  Desc: string;
  IssuePurpose: string;
  Number: string;
  RevisionDate: string;
  RevisionNumber: string;
  Title: string;
  customProperties: CustomProp[];
  layoutInformation: LayoutProp[];
};

const App: React.FC<{ children?: React.ReactNode }> = () => {
  const [xmlData, setXmlData] = React.useState<any | null>(null);
  const [originalFile, setOriginalFile] = React.useState<File | null>(null);
  const [sheets, setSheets] = React.useState<Sheet[]>([]);
  const sswrapRef = React.useRef<HTMLDivElement>(null);

  const downloadXlsx = async () => {
    const customProps = sheets[0]?.customProperties.map((v) => v.name);
    const layoutProps = ["FileName", "Name", "Relative_FileName"];
    const layoutPropsNames = ["File", "Layout Name", "Relative File Path"];

    const data: SheetData = [
      [
        "Title",
        "Number",
        "Description",
        "Revision Number",
        "Revision Date",
        "Issue Purpose",
        "Category",
        ...customProps,
        ...layoutPropsNames,
      ].map((v) => ({ value: v, fontWeight: "bold" })),
    ];
    sheets.forEach((sheet) => {
      data.push([
        { type: String, value: sheet.Title },
        { type: String, value: sheet.Number },
        { type: String, value: sheet.Desc },
        { type: String, value: sheet.RevisionNumber },
        { type: String, value: sheet.RevisionDate },
        { type: String, value: sheet.IssuePurpose },
        { type: String, value: sheet.Category },
        ...customProps.map((name) => ({
          type: String,
          value: sheet.customProperties.find((v) => v.name === name)?.value || "",
        })),
        ...layoutProps.map((name) => ({
          type: String,
          value: sheet.layoutInformation.find((v) => v.name === name)?.value || "",
        })),
      ]);
    });
    const blob = await writeXlsxFile(data, {
      // columns: [{width: 100 }], // (optional) column widths, etc.
    });
    const downloadLink = document.createElement("a");
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = `${originalFile?.name?.replace("dst", "xlsx") || `Sheetset.xlsx`}`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
  };
  const download = () => {
    const data = JSON.parse(JSON.stringify(xmlData));
    const xmlSheets = findObjects(data, "AcSmSheet", "_ANY").flatMap((v) => v) as any as XMLSheet[];

    xmlSheets.map((xmlsheet, si) => {
      const sheet = sheets[si];
      xmlsheet.AcSmProp = [
        { _attributes: { propname: "Category", vt: "8" }, _text: sheet.Category },
        { _attributes: { propname: "Desc", vt: "8" }, _text: sheet.Desc },
        { _attributes: { propname: "IssuePurpose", vt: "8" }, _text: sheet.IssuePurpose },
        { _attributes: { propname: "Number", vt: "8" }, _text: sheet.Number },
        { _attributes: { propname: "RevisionDate", vt: "8" }, _text: sheet.RevisionDate },
        { _attributes: { propname: "RevisionNumber", vt: "8" }, _text: sheet.RevisionNumber },
        { _attributes: { propname: "Title", vt: "8" }, _text: sheet.Title },
      ].filter((prop) => prop._text !== "" || typeof prop._text !== "undefined") as XMLSheetProp[];
      if (xmlsheet.AcSmProp.length === 1) {
        xmlsheet.AcSmProp = xmlsheet.AcSmProp[0];
      }
      // sheet.Title
      const customProps = ((Array.isArray(xmlsheet.AcSmCustomPropertyBag)
        ? xmlsheet.AcSmCustomPropertyBag
        : [xmlsheet.AcSmCustomPropertyBag]) || []) as XMLCustomProp[];
      customProps.forEach((_, i) => {
        const name = customProps[i].AcSmCustomPropertyValue._attributes.propname;
        const root = customProps[i]?.AcSmCustomPropertyValue?.AcSmProp?.find((k) => k._attributes.propname === "Value");
        const newValue = sheets[si].customProperties.find((prop) => prop.name === name)?.value;
        if (root) root["_text"] = typeof newValue === "string" ? newValue : "";
      });
      return xmlsheet;
    });

    const back2xml = js2xml(data as any, { compact: true });
    const textEncoder = new TextEncoder();
    const backView = new DataView(textEncoder.encode(back2xml).buffer);
    for (let i = 0; i < backView.byteLength; i++) {
      let byteValue = backView.getUint8(i);
      let flippedValue = flipBits(byteValue, true);
      backView.setUint8(i, flippedValue);
    }

    const blob = new Blob([backView.buffer], { type: "application/octet-stream" });
    const downloadLink = document.createElement("a");
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = `${originalFile?.name || `Sheetset.dst`}`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
  };

  const onDrop = React.useCallback((acceptedFiles: File[]) => {
    acceptedFiles.forEach((file) => {
      if (file.name.split(".").reverse()[0].toLowerCase() !== "dst") return;

      setOriginalFile(file);
      const reader = new FileReader();

      reader.onabort = () => console.log("file reading was aborted");
      reader.onerror = () => console.log("file reading has failed");
      reader.onload = () => {
        const view = new DataView(reader.result as ArrayBuffer);
        for (let i = 0; i < view.byteLength; i++) {
          let byteValue = view.getUint8(i);
          let flippedValue = flipBits(byteValue);
          view.setUint8(i, flippedValue);
        }
        const textDecoder = new TextDecoder("utf-8"); // Specify the encoding, e.g., 'utf-8'
        const data = xml2js(textDecoder.decode(view), { compact: true });

        const spreadsheets = findObjects(data, "AcSmSheet", "_ANY").flatMap((v) => v) as any as XMLSheet[];
        const sheetData: Sheet[] = spreadsheets.map((sheet) => {
          const sheetProps = ((Array.isArray(sheet.AcSmProp) ? sheet.AcSmProp : [sheet.AcSmProp]) ||
            []) as XMLSheetProp[];
          return {
            root: sheet,
            Title: sheetProps.find((prop) => prop._attributes.propname == "Title")?._text || "",
            Desc: sheetProps.find((prop) => prop._attributes.propname == "Desc")?._text || "",
            Number: sheetProps.find((prop) => prop._attributes.propname == "Number")?._text || "",
            Category: sheetProps.find((prop) => prop._attributes.propname == "Category")?._text || "",
            RevisionNumber: sheetProps.find((prop) => prop._attributes.propname == "RevisionNumber")?._text || "",
            RevisionDate: sheetProps.find((prop) => prop._attributes.propname == "RevisionDate")?._text || "",
            IssuePurpose: sheetProps.find((prop) => prop._attributes.propname == "IssuePurpose")?._text || "",
            customProperties: (
              ((Array.isArray(sheet.AcSmCustomPropertyBag)
                ? sheet.AcSmCustomPropertyBag
                : [sheet.AcSmCustomPropertyBag]) || []) as XMLCustomProp[]
            ).map((prop) => {
              const name = prop.AcSmCustomPropertyValue._attributes.propname;
              const root = prop?.AcSmCustomPropertyValue?.AcSmProp?.find((k) => k._attributes.propname === "Value");
              const key = "_text";
              const value = root ? root[key] : "";
              return {
                root: root,
                name,
                key,
                value,
              } as CustomProp;
            }),
            layoutInformation: (
              ((Array.isArray(sheet.AcSmAcDbLayoutReference.AcSmProp)
                ? sheet.AcSmAcDbLayoutReference.AcSmProp
                : [sheet.AcSmAcDbLayoutReference.AcSmProp]) || []) as XMLLayoutProp[]
            ).map((prop) => {
              const name = prop._attributes.propname;
              const root = prop;
              const key = "_text";
              const value = root ? root[key] : "";
              return {
                root,
                name,
                key,
                value,
              } as LayoutProp;
            }),
          };
        });
        setSheets(sheetData);
        setXmlData(data);
      };
      reader.readAsArrayBuffer(file);
    });
  }, []);
  const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop });

  const customProps = sheets[0]?.customProperties.map((v) => v.name);
  const layoutProps = ["FileName", "Name", "Relative_FileName"];
  const layoutPropsNames = ["File", "Layout Name", "Relative File Path"];

  return (
    <Wrapper>
      <SideWrapper>
        <h1 style={{ textAlign: "center" }}>DST Sheetset properties editor </h1>
        {/* {xmlData && <pre style={{ maxWidth: "80vw" }}>{JSON.stringify(xmlData, 1 as any, 1)}</pre>} */}
        <div style={{ maxWidth: 460, marginBottom: 20 }}>
          <h3>How to use this tool</h3>
          <ol>
            <li>
              Slect Sheetset (dst) file.
              <br />
              (file isn't uploaded and is processed locally)
            </li>
            <li>
              Update any of the sheets variables
              <br />
              (Title, Number, Custom Fields,..)
            </li>
            <li>Download updated and save in the same location with the same name.</li>
          </ol>
          <p>
            Notes:
            <ol>
              <li>
                Keep your previous Sheetset (.dst) file around as a backup as this tool doesn't guarantee that the
                updated dst file will work as intended. This tool is made available as is, and doesn't provide any
                guarantees.
              </li>
              <li>
                This tool intentionally doesn't provide functionality to change sheetset file to drawing links since
                it's a two way process (dst links to dwg, dwg links to dst). Allowing such functionality would create
                only one way linkage.
              </li>
            </ol>
          </p>
        </div>
        <Dropzone {...getRootProps()} $dragActive={isDragActive}>
          <input {...getInputProps()} />
          <p>Drag 'n' drop your sheetset.dst file here, or click to select files</p>
        </Dropzone>
        {xmlData && (
          <>
            <Download
              onClick={() => {
                downloadXlsx();
              }}
            >
              Download as Excel file
            </Download>
            <Download
              onClick={() => {
                download();
              }}
            >
              Download sheetset file
            </Download>
          </>
        )}

        <p style={{ textAlign: "center", marginTop: "auto" }}>
          Made with love by <a href="https://github.com/abrman">Matthew Abrman</a>
          <br />
          Check the <a href="https://github.com/abrman/dst-edit">github repo</a> submit any issues!
        </p>
      </SideWrapper>
      <SSWrapper ref={sswrapRef}>
        {xmlData && (
          <Spreadsheet
            // hideColumnIndicators={true}
            onChange={(data) => {
              data = data.filter((_, i) => i < sheets.length);
              // const stringifyToDepth = (obj: object, depth: number) => {
              //   const seen = new Set();

              //   return JSON.stringify(obj, function (key, value) {
              //     if (depth !== undefined) {
              //       if (seen.has(value) || typeof value === "function") {
              //         return;
              //       }

              //       seen.add(value);

              //       if (depth > 0) {
              //         depth--;
              //         return value;
              //       }
              //     }

              //     return value;
              //   });
              // };
              setSheets((prev) => {
                const newSheets = JSON.parse(JSON.stringify(prev)) as Sheet[];
                // const prevJSON = stringifyToDepth(prev, 3);
                data.forEach((col, i) => {
                  newSheets[i].Title = typeof col[0]?.value === "string" ? col[0]?.value.slice(0, 64) : "";
                  newSheets[i].Number = typeof col[1]?.value === "string" ? col[1]?.value : "";
                  newSheets[i].Desc = typeof col[2]?.value === "string" ? col[2]?.value : "";
                  newSheets[i].RevisionNumber = typeof col[3]?.value === "string" ? col[3]?.value : "";
                  newSheets[i].RevisionDate = typeof col[4]?.value === "string" ? col[4]?.value : "";
                  newSheets[i].IssuePurpose = typeof col[5]?.value === "string" ? col[5]?.value : "";
                  newSheets[i].Category = typeof col[6]?.value === "string" ? col[6]?.value : "";
                  newSheets[i].customProperties = customProps
                    .map((customPropName, j) => {
                      const updatedCustomProp = newSheets[i].customProperties.find((v) => v.name === customPropName);
                      if (!updatedCustomProp) return false;
                      (updatedCustomProp as any).value = typeof col[j + 7]?.value === "string" ? col[j + 7]?.value : "";
                      return { ...updatedCustomProp };
                    })
                    .filter((v) => v !== false) as CustomProp[];

                  if (col[0]?.value && col[0]?.value.length > 64) {
                    toast(
                      `Title on figure ${i + 1} (${
                        newSheets[i].Number
                      }) has been trimmed to follow the max limit of 64 characters`
                    );
                  }
                });

                // if (stringifyToDepth(newSheets, 3) === prevJSON) return prev;
                return newSheets;
              });
            }}
            columnLabels={[
              "Title",
              "Number",
              "Description",
              "Revision Number",
              "Revision Date",
              "Issue Purpose",
              "Category",
              ...customProps,
              ...layoutPropsNames,
            ]}
            data={sheets.map((sheet) => {
              return [
                { value: sheet.Title },
                { value: sheet.Number },
                { value: sheet.Desc },
                { value: sheet.RevisionNumber },
                { value: sheet.RevisionDate },
                { value: sheet.IssuePurpose },
                { value: sheet.Category },
                ...customProps.map((name) => ({
                  value: sheet.customProperties.find((v) => v.name === name)?.value || "",
                })),
                ...layoutProps.map((name) => ({
                  value: sheet.layoutInformation.find((v) => v.name === name)?.value || "",
                  readOnly: true,
                })),
              ];
            })}
          />
        )}
      </SSWrapper>
    </Wrapper>
  );
};

export default App;
