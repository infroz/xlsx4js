export { createWorkbook, createWorksheet } from "./lib/index.js";
import { createWorkbook, createWorksheet } from "./lib/index.js";

const workbook = createWorkbook({ name: "My Workbook" });

const sheeet = createWorksheet({ name: "My Sheet", rows: [{ test: 123, string: "My Value1"}, { test: 456, string: "My Second Value1"}] });

workbook.addSheet(sheeet);

workbook.write({ file: { filename: "MyWorkbook.xlsx" } });

