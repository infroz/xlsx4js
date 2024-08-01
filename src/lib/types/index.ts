/**
 * Valid data types for a cell
 */
export type ValidDataType =
  | string
  | number
  | boolean
  | Date
  | Error
  | null
  | undefined;

/**
 * Cell type Boolean OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellTypeBoolean = "b";

/**
 * Cell type Date OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellTypeDate = "d";

/**
 * Cell type Error OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellTypeError = "e";

/**
 * Cell type Number OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellTypeNumber = "n";

/**
 * Cell type SharedString OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellTypeSharedString = "s";

/**
 * Cell type String OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellTypeString = "str";

/**
 * Cell type InlineString OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellTypeInlineString = "inlineStr";

/**
 * Cell type Empty OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellTypeEmpty = "z";

/**
 * Cell type OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
export type CellType =
  | CellTypeBoolean
  | CellTypeDate
  | CellTypeError
  | CellTypeNumber
  | CellTypeSharedString
  | CellTypeString
  | CellTypeInlineString
  | CellTypeEmpty;

export const isValidType = (value: any): boolean => {
  if (typeof value === "string") return true;
  if (typeof value === "number") return true;
  if (typeof value === "boolean") return true;
  if (value instanceof Date) return true;
  if (value instanceof Error) return true;
  if (value === null || value === undefined) return true;

  return false;
};

/**
 * Get the cell type
 * @param value - The value of the cell
 * @returns CellType
 */
export const getCellType = (value: ValidDataType): CellType => {
  if (typeof value === "string") return "str";
  if (typeof value === "number") return "n";
  if (typeof value === "boolean") return "b";
  if (value instanceof Date) return "d";
  if (value instanceof Error) return "e";
  if (value === null || value === undefined) return "z";
  throw new Error(`Invalid cell type: ${value}`);
};

/**
 * Keep track of the rows in the worksheet
 */
export type Row = { [key: string]: ValidDataType };
