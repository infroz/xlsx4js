import { Sheet, Workbook } from "./lib";

export default function main() {
    console.log('Hello, world!'); 
    const workbook = new Workbook();
    const sheet1 = new Sheet('My custom sheet');
    workbook.addSheet(sheet1);

    workbook.write();
}

main();