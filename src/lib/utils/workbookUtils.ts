import { Worksheet } from "../createWorksheet.js";
import { getCellType, Row } from "../index.js";

export const getSheetOverride = (sheet: Worksheet) => `<Override PartName="/xl/worksheets/${sheet.getName()}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
export const getSheetRels = (sheet: Worksheet) => `<Relationship Id="${sheet.getId()}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheet.getName()}.xml"/>`;
export const getSheets = (sheet: Worksheet) => `<sheet name="${sheet.getName()}" sheetId="${2}" r:id="${sheet.getId()}"/>`;

const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");

const renderRow = (row: Row, index: number) => `<row r="${index + 1}">
        ${Object.entries(row).map(([_, value], i) => `<c r="${alphabet[i]}${index + 1}" t="${getCellType(value)}"><v>${value}</v></c>`).join("\n")}
    </row>`;

const renderCell = (value: any, index: number) => `<c r="${alphabet[index]}1" t="${getCellType(value)}"><v>${value}</v></c>`;

export const getSheetFile = (sheet: Worksheet) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
         ${sheet.rows.map(renderRow).join("\n")} 
      </sheetData>
  </worksheet>`;
