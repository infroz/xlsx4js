import createWorkbook from "./lib/createWorkbook.js";
import createWorksheet from "./lib/createWorksheet.js";

export { createWorkbook, createWorksheet } from "./lib/index.js";

const book = createWorkbook({ name: "My Sheet" });
const sheet = createWorksheet([
  { a: 1, b: "one" },
  { a: 2, b: "two" },
  { a: 3, b: "three" },
]);

book.addSheet(sheet);

book.write();
