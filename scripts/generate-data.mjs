import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const books = [
  { workbook: "Entregable 1 · Base de datos", file: "Planning_Carga_Trabajo_Entregable_1(1).xlsx" },
  { workbook: "Planning generado por IA", file: "Planning_IA(1).xlsx" },
  { workbook: "Entregable 3 · Dashboard", file: "Entregable_3_Dashboard_Indicadores(1).xlsx" },
  { workbook: "Entregable 2 · Planning", file: "Entregable_2_Planning_Carga_Trabajo(1).xlsx" }
];

const clean = value => String(value ?? "")
  .normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();

function parseSheet(ws, sheet) {
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
  if (!matrix.length) return [];
  let headerIndex = 0;
  if (clean(sheet) !== "resumenes") {
    let best = -Infinity;
    for (let i = 0; i < Math.min(15, matrix.length); i++) {
      const values = (matrix[i] ?? []).map(value => String(value).trim()).filter(Boolean);
      const score = values.length * 3 + new Set(values.map(clean)).size - (values.length <= 1 ? 8 : 0);
      if (score > best) { best = score; headerIndex = i; }
    }
  }
  const rawHeaders = (matrix[headerIndex] ?? [])
    .map((value, index) => String(value).trim() || `Columna ${index + 1}`);
  const seen = new Map();
  const headers = rawHeaders.map(header => {
    const n = (seen.get(header) ?? 0) + 1;
    seen.set(header, n);
    return n === 1 ? header : `${header} ${n}`;
  });
  return matrix.slice(headerIndex + 1)
    .filter(row => row.some(value => String(value ?? "").trim()))
    .map(row => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
}

const dataDir = path.join(process.cwd(), "data");
const tables = books.flatMap(book => {
  const workbook = XLSX.readFile(path.join(dataDir, book.file));
  return workbook.SheetNames.map(sheet => {
    const rows = parseSheet(workbook.Sheets[sheet], sheet);
    return {
      workbook: book.workbook,
      file: book.file,
      sheet,
      columns: Object.keys(rows[0] ?? {}),
      rows
    };
  }).filter(table => table.rows.length && table.columns.length);
});

fs.writeFileSync(path.join(dataDir, "catalog.json"), JSON.stringify(tables));
console.log(`Generado catalog.json con ${tables.length} hojas y ${tables.reduce((n, table) => n + table.rows.length, 0)} registros.`);
