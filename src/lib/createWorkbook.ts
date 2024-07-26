import JSZip from "jszip";
import { Worksheet } from "./createWorksheet.js";
import {
  getSheetFile,
  getSheetOverride,
  getSheetRels,
  getSheets,
} from "./utils/workbookUtils.js";
import fs from "fs";

type Options = {
  name: string;
  sheets?: Worksheet[]; // todo: implement sheet
};

export type Workbook = {
  addSheet: (sheet: any) => void;
  addSheets: (sheets: any[]) => void;
  write: (options?: {
    file?: {
      filename?: string;
      extension?: "xlsx";
    };
  }) => void;
};

const ensureUniqueSheetName = (sheets: Worksheet[], name: string) => {
  let count = 0;
  let newName = name;
  while (sheets.some((sheet) => sheet.getName() === newName)) {
    count++;
    newName = `${name} ${count}`;
  }
  return newName;
};

export const createWorkbook = (options?: Options): Workbook => {
  let _name = options?.name ?? "Workbook";
  let _sheets: Worksheet[] = options?.sheets ?? [];

  return {
    addSheet: (sheet) => {
      if (_sheets.some((s) => s.getName() === sheet.getName())) {
        sheet.setName(ensureUniqueSheetName(_sheets, sheet.getName()));
      }
      _sheets.push(sheet);
    },
    addSheets: (sheets) => {
      if (
        sheets.some((sheet) =>
          _sheets.some((s) => s.getName() === sheet.getName())
        )
      ) {
        sheets.forEach((sheet) => {
          sheet.setName(ensureUniqueSheetName(_sheets, sheet.getName()));
        });
      }
      _sheets = [...sheets, ..._sheets];
    },
    write: (options) => {
      const zip = new JSZip();

      // [Content_Types].xml
      const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      ${_sheets.map(getSheetOverride)}
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  </Types>`;
      zip.file("[Content_Types].xml", contentTypes);

      // _rels/.rels
      const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  </Relationships>`;
      zip.file("_rels/.rels", rels);

      // xl/_rels/workbook.xml.rels
      const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      ${_sheets.map(getSheetRels)}
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`;
      zip.file("xl/_rels/workbook.xml.rels", workbookRels);
      console.log("\nworkbookRels:\n", workbookRels);

      // xl/workbook.xml
      const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets>
          ${_sheets.map(getSheets)}
      </sheets>
  </workbook>`;
      zip.file("xl/workbook.xml", workbook);

      _sheets.forEach((sheet) => {
        zip.file(`xl/worksheets/${sheet.getName()}.xml`, getSheetFile(sheet));
      });

      // generate file
      zip
        .generateNodeStream({ type: "nodebuffer", streamFiles: true })
        .pipe(fs.createWriteStream(`${options?.file?.filename ?? _name}.xlsx`));
    },
  };
};

export default createWorkbook;
