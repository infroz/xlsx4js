import { Row, isValidType } from "./types/index.js";

type BaseOptions = { name: string };

type Options = BaseOptions &
  ({ rows: Row[]; data: never } | { data: object[] | Row[]; rows: never });

export type Worksheet = {
  rows: Row[];
  addRow: (row: Row) => void;
  addRows: (rows: Row[]) => void;
  getName: () => string;
  setName: (name: string) => void;
  getData: () => Row[];
  getId: () => string;
  setId: (id: string) => void;
};

const genericObjectsToRows = (
  arr?: { [key: string]: any }[],
): Row[] | undefined => {
  return arr?.map((element) => {
    const newElement: Row = {};

    for (let key in element)
      if (element.hasOwnProperty(key))
        newElement[key] = isValidType(element[key])
          ? element[key]
          : JSON.stringify(element[key]);
    return newElement;
  });
};

export const createWorksheet = <TData extends object[] | Row[]>(
  data: TData,
  options?: Options,
): Worksheet => {
  // Worksheet state
  let name = options?.name ?? "Sheet";

  const rows: Row[] = genericObjectsToRows(data) ?? [];
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
    setName: (newName) => {
      name = newName;
    },
    getData: () => rows,
    getId: () => id,
    setId: (newId) => {
      id = newId;
    },
  };
};

export default createWorksheet;
