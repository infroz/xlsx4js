import createWorkbook from "./lib/createWorkbook.js";
import createWorksheet from "./lib/createWorksheet.js";

export * from "./lib/createWorkbook.js";
export * from "./lib/createWorksheet.js";

const workbook = createWorkbook({ name: "My Workbook" });

const sheeet = createWorksheet({ name: "My Sheet", rows: [{ test: 123, string: "My Value1"}, { test: 456, string: "My Second Value1"}] });

workbook.addSheet(sheeet);

workbook.write({ file: { filename: "MyWorkbook.xlsx" } });

