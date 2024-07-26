
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
    setName: (name: string) => void;
    getData: () => Row[];
    getId: () => string;
    setId: (id: string) => void;
}

export const createWorksheet = (options?: Options): Worksheet => {
    // Worksheet state
    let name = options?.name ?? "Sheet";
    const rows = options?.rows ?? [];
    let id: string;

    return {
        rows: rows,
        addRow: (row) => {
            rows.push(row);
        },
        addRows: (rows) => {
            rows.push(...rows);
        },
        getName: () => name,
        setName: (newName) => { name = newName; },
        getData: () => rows,
        getId: () => id,
        setId: (newId) => {
            id = newId;
        }
    };
};

export default createWorksheet;