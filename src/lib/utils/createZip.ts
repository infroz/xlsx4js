import JSZip from "jszip";
import { Worksheet } from "../createWorksheet.js";
import {
  getSheetFile,
  getSheetOverride,
  getSheetRels,
  getSheets,
} from "./workbookUtils.js";

export const createZip = ({ sheets }: { sheets: Worksheet[] }): JSZip => {
  const zip = new JSZip();

  // [Content_Types].xml
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      ${sheets.map(getSheetOverride)}
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
      ${sheets.map(getSheetRels)}
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`;
  zip.file("xl/_rels/workbook.xml.rels", workbookRels);

  // xl/workbook.xml
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets>
          ${sheets.map(getSheets)}
      </sheets>
  </workbook>`;
  zip.file("xl/workbook.xml", workbook);

  sheets.forEach((sheet) => {
    zip.file(`xl/worksheets/${sheet.getName()}.xml`, getSheetFile(sheet));
  });

  return zip;
};
