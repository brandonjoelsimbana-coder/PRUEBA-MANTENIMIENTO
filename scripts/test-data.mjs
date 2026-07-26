import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const dir = path.join(process.cwd(), "data");
const files = fs.readdirSync(dir).filter(file => file.endsWith(".xlsx"));
if (files.length !== 4) throw new Error(`Se esperaban 4 Excel y se encontraron ${files.length}.`);
let sheets = 0;
for (const file of files) {
  const workbook = XLSX.readFile(path.join(dir, file));
  if (!workbook.SheetNames.length) throw new Error(`${file} no contiene hojas.`);
  sheets += workbook.SheetNames.length;
}
console.log(`OK: ${files.length} archivos y ${sheets} hojas disponibles.`);
