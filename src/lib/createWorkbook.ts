import JSZip from "jszip";
import { Worksheet } from "./createWorksheet.js";
import {
  getSheetFile,
  getSheetOverride,
  getSheetRels,
  getSheets,
} from "./utils/workbookUtils.js";
import fs from "fs";
import { createZip } from "./utils/createZip.js";

type Options = {
  name: string;
  sheets?: Worksheet[]; // todo: implement sheet
};

export type Workbook = {
  addSheet: (sheet: any) => void;
  addSheets: (sheets: any[]) => void;
  write: () => void;
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
          _sheets.some((s) => s.getName() === sheet.getName()),
        )
      ) {
        sheets.forEach((sheet) => {
          sheet.setName(ensureUniqueSheetName(_sheets, sheet.getName()));
        });
      }
      _sheets = [...sheets, ..._sheets];
    },
    write: () => {
      const zip = createZip({ sheets: _sheets });

      // generate file
      zip
        .generateNodeStream({ type: "nodebuffer", streamFiles: true })
        .pipe(fs.createWriteStream(`${_name}.xlsx`));
    },
  };
};

export default createWorkbook;
