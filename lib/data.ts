import fs from "node:fs";
import path from "node:path";
import * as XLSX from "xlsx";
import type { DataTable, Row } from "./types";

const files = [
  { name: "Entregable 1 · Base de datos", file: "Planning_Carga_Trabajo_Entregable_1(1).xlsx" },
  { name: "Planning generado por IA", file: "Planning_IA(1).xlsx" },
  { name: "Entregable 3 · Dashboard", file: "Entregable_3_Dashboard_Indicadores(1).xlsx" },
  { name: "Entregable 2 · Planning", file: "Entregable_2_Planning_Carga_Trabajo(1).xlsx" }
];

const clean = (value: unknown) =>
  String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();

function parseSheet(ws: XLSX.WorkSheet, sheet: string): Row[] {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "", raw: false });
  if (!matrix.length) return [];
  let headerIndex = 0;
  if (clean(sheet) !== "resumenes") {
    let best = -Infinity;
    for (let i = 0; i < Math.min(15, matrix.length); i++) {
      const values = (matrix[i] ?? []).map(v => String(v).trim()).filter(Boolean);
      const score = values.length * 3 + new Set(values.map(clean)).size - (values.length <= 1 ? 8 : 0);
      if (score > best) { best = score; headerIndex = i; }
    }
  }
  const rawHeaders = (matrix[headerIndex] ?? []).map((v, i) => String(v).trim() || `Columna ${i + 1}`);
  const seen = new Map<string, number>();
  const headers = rawHeaders.map(header => {
    const n = (seen.get(header) ?? 0) + 1;
    seen.set(header, n);
    return n === 1 ? header : `${header} ${n}`;
  });
  return matrix.slice(headerIndex + 1)
    .filter(row => row.some(value => String(value ?? "").trim()))
    .map(row => Object.fromEntries(headers.map((header, i) => [header, (row[i] ?? "") as Row[string]])));
}

let cache: DataTable[] | null = null;

export function loadTables(): DataTable[] {
  if (cache) return cache;
  const dataDir = path.join(process.cwd(), "data");
  cache = files.flatMap(book => {
    const workbook = XLSX.readFile(path.join(dataDir, book.file), { cellDates: true });
    return workbook.SheetNames.map(sheet => {
      const rows = parseSheet(workbook.Sheets[sheet], sheet);
      return { workbook: book.name, file: book.file, sheet, rows, columns: Object.keys(rows[0] ?? {}) };
    }).filter(table => table.rows.length && table.columns.length);
  });
  return cache;
}

export function buildCatalog(tables: DataTable[]) {
  return tables.map(table => ({
    workbook: table.workbook,
    sheet: table.sheet,
    rows: table.rows.length,
    columns: table.columns.map(column => {
      const values = [...new Set(table.rows.map(row => String(row[column] ?? "").trim()).filter(Boolean))];
      const numeric = values.length > 0 && values.filter(value => Number.isFinite(Number(value.replace(",", ".")))).length / values.length > 0.7;
      return {
        name: column,
        type: numeric ? "number" : "text",
        examples: numeric ? values.slice(0, 4) : values.filter(value => value.length < 60).slice(0, 18)
      };
    })
  }));
}
