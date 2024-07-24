import JSZip from "jszip";
import fs from "fs";
import { getSheetFile, getSheetOverride, getSheetRels, getSheets } from "./utils/workbookUtils.js";
import Worksheet from "./Worksheet.js";

/** @deprecated */
export default class Workbook {
  _sheets: Worksheet[] = [];
  addSheet(sheet: Worksheet) {
    this._sheets.push(new Worksheet({ 
        id: `sheetId${this._sheets.length + 1}`,
        name: sheet.getName(),
        relationshipId: `sheetRId${this._sheets.length + 1}`,
        rows: sheet.getData().rows
    }));
  }
  write() {
    const zip = new JSZip();

    // [Content_Types].xml
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      ${this._sheets.map(getSheetOverride)}
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  </Types>`;
    zip.file("[Content_Types].xml", contentTypes);
    console.log("\ncontentTypes:\n", contentTypes);

    // _rels/.rels
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  </Relationships>`;
    zip.file("_rels/.rels", rels);

    // xl/_rels/workbook.xml.rels
    const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      ${this._sheets.map(getSheetRels)}
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`;
    zip.file("xl/_rels/workbook.xml.rels", workbookRels);
    console.log("\nworkbookRels:\n", workbookRels);

    // xl/workbook.xml
    const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets>
          <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
          ${this._sheets.map(getSheets)}
      </sheets>
  </workbook>`;
    zip.file("xl/workbook.xml", workbook);
    console.log("\nworkbook:\n", workbook);

    // xl/worksheets/sheet1.xml
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
          <row r="1">
              <c r="A1"><v>0</v></c>
              <c r="B1"><v>1</v></c>
          </row>
          <row r="2">
              <c r="A2"><v>2</v></c>
              <c r="B2"><v>3</v></c>
          </row>
      </sheetData>
  </worksheet>`;
    zip.file("xl/worksheets/sheet1.xml", sheet);
    console.log("\nsheet:\n", sheet);

    this._sheets.forEach((sheet) => {
        zip.file(`xl/worksheets/${sheet.getName()}.xml`, getSheetFile(sheet));
        console.log(`\n${sheet.getName()}:\n`, getSheetFile(sheet));
    });

    // xl/sharedStrings.xml
    const sharedStrings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
      <si><t>Header1</t></si>
      <si><t>Header2</t></si>
      <si><t>Data1</t></si>
      <si><t>Data2</t></si>
  </sst>`;
    // zip.file("xl/sharedStrings.xml", sharedStrings);

    // xl/styles.xml
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1">
          <font>
              <sz val="11"/>
              <color theme="1"/>
              <name val="Calibri"/>
              <family val="2"/>
              <scheme val="minor"/>
          </font>
      </fonts>
      <fills count="2">
          <fill>
              <patternFill patternType="none"/>
          </fill>
          <fill>
              <patternFill patternType="gray125"/>
          </fill>
      </fills>
      <borders count="1">
          <border>
              <left/>
              <right/>
              <top/>
              <bottom/>
              <diagonal/>
          </border>
      </borders>
      <cellStyleXfs count="1">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
      </cellStyleXfs>
      <cellXfs count="1">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
      </cellXfs>
  </styleSheet>`;
    // zip.file("xl/styles.xml", styles);

    // generate file
    zip
      .generateNodeStream({ type: "nodebuffer", streamFiles: true })
      .pipe(fs.createWriteStream("workbook.xlsx"))
      .on("finish", function () {
        console.log("workbook.xlsx written.");
      });
  }
  writeBlob () {
    throw new Error("Method not implemented.");
  };
  writeBlobAsync () {
    throw new Error("Method not implemented.");
  }
  writeBuffer () {
    throw new Error("Method not implemented.");
  }
  writeBufferAsync () {
    throw new Error("Method not implemented.");
  }
}
