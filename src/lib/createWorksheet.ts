import { Row } from "./types/index.js";

type Options = {
    name: string;
    rows: Row[];
}

export type Worksheet = {
    rows: Row[];
    addRow: (row: Row) => void;
    addRows: (rows: Row[]) => void;
    getName: () => string;
    getData: () => Row[];
    getId: () => string;
    setId: (id: string) => void;
}

export const createWorksheet = (options?: Options): Worksheet => {
    // Worksheet state
    const name = options?.name ?? "Sheet";
    const rows = options?.rows ?? [];
    let id: string;

    return {
        rows: rows,
        addRow: (row) => {
            console.log("Adding row to worksheet", row);
        },
        addRows: (rows) => {
            console.log("Adding rows to worksheet", rows);
        },
        getName: () => name,
        getData: () => rows,
        getId: () => id,
        setId: (newId) => {
            id = newId;
        }
    };
};

export default createWorksheet;