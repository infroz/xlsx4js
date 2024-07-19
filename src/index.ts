import { Workbook, Worksheet } from "./lib";

export default function main() {
    console.log('Hello, world!'); 
    const workbook = new Workbook();
    const sheet1 = new Worksheet({ name: "Custom1", id: "", relationshipId: "", rows: [
        { "A": "Hello", "B": "World" },
        { "A": 1, "B": 2 }
    ] });
    workbook.addSheet(sheet1);

    workbook.write();
}

main();