

/**
 * Internal type for handling worksheet data 
 * 
 * @param name - The name of the worksheet
 * @param id - The id of the worksheet
 * @param relationshipId - The relationship id of the worksheet - not in use
 * @param rows - The rows of the worksheet
 */
type Sheet = {
    name: string;
    id: string;
    relationshipId: string;
    rows: Row[];
}

/**
 * Valid data types for a cell
 */
export type ValidDataType = string | number | boolean | Date | Error | null | undefined;

/**
 * Cell type Boolean OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellTypeBoolean = "b";

/**
 * Cell type Date OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellTypeDate = "d";

/**
 * Cell type Error OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellTypeError = "e";

/**
 * Cell type Number OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellTypeNumber = "n";

/**
 * Cell type SharedString OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellTypeSharedString = "s";

/**
 * Cell type String OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellTypeString = "str";

/**
 * Cell type InlineString OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellTypeInlineString = "inlineStr";

/**
 * Cell type Empty OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellTypeEmpty = "z";

/**
 * Cell type OpenXML-3.0.1
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1#fields
 */
type CellType = CellTypeBoolean | CellTypeDate | CellTypeError | CellTypeNumber | CellTypeSharedString | CellTypeString | CellTypeInlineString | CellTypeEmpty;

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
}

/**
 * Keep track of the rows in the worksheet
 */
export type Row = { [key: string]: ValidDataType }; 

/**
 * @deprecated
 * class object for creating and managing worksheets
 * 
 * A workbook contains one or more worksheets
 */
export default class Worksheet {
    #data: Sheet;
    constructor(data: Sheet) {
       this.#data = data; 
    }

    getName () { return this.#data.name; }
    setName (name: string) { this.#data.name = name; }
    getData () { return this.#data; }
}