import { ISheet } from "./ISheet";

export interface IWorkbook {
    addSheet: (sheet: ISheet) => void;
    write: () => void;
};