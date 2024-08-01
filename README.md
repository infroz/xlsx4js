[![PNPM INSTALL & BUILD](https://github.com/infroz/xlsx4js/actions/workflows/pnpm-ci.yml/badge.svg)](https://github.com/infroz/xlsx4js/actions/workflows/pnpm-ci.yml) [![Node.js Package](https://github.com/infroz/xlsx4js/actions/workflows/npm-publish.yml/badge.svg)](https://github.com/infroz/xlsx4js/actions/workflows/npm-publish.yml)

# XLSX4JS

This project is in its early stages of development, if you have find bugs or have suggestions please
create an issue or fork the project to create pull-reqest.

# How to use

```ts
// Create a workbook
const workbook = createWorkbook({ name: "My Workbook" });

// Create a sheet
const sheeet = createWorksheet(
  [
    { someInteger: 123, someString: "My Value1" },
    { someInteger: 456, someString: "My Second Value1" },
  ]
);

// Add the sheet to your book
workbook.addSheet(sheeet);

// Write the file, currently using node's fs
workbook.write({ file: { filename: "MyWorkbook" } });
```
