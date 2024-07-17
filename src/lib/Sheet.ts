import { ISheet } from "../types/ISheet";

export class Sheet implements ISheet {
    name: string;

    constructor(name: string) {
        this.name = name;
    }

}