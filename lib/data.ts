import embeddedTables from "../data/catalog.json";
import type { DataTable } from "./types";

const tables = embeddedTables as DataTable[];

export function loadTables(): DataTable[] {
  return tables;
}

export function buildCatalog(dataTables: DataTable[]) {
  return dataTables.map(table => ({
    workbook: table.workbook,
    sheet: table.sheet,
    rows: table.rows.length,
    columns: table.columns.map(column => {
      const values = [...new Set(table.rows.map(row => String(row[column] ?? "").trim()).filter(Boolean))];
      const numeric = values.length > 0 &&
        values.filter(value => Number.isFinite(Number(value.replace(",", ".")))).length / values.length > 0.7;
      return {
        name: column,
        type: numeric ? "number" : "text",
        examples: numeric
          ? values.slice(0, 4)
          : values.filter(value => value.length < 60).slice(0, 18)
      };
    })
  }));
}
