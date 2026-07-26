import type { DataTable, QueryPlan, QueryResult, Row } from "./types";

const clean = (value: unknown) =>
  String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
const tokens = (value: string) => clean(value).split(/[^a-z0-9]+/).filter(Boolean);
const singular = (word: string) => word.endsWith("es") && word.length > 4 ? word.slice(0, -2) : word.endsWith("s") && word.length > 3 ? word.slice(0, -1) : word;
const fmt = (value: number) => new Intl.NumberFormat("es-EC", { maximumFractionDigits: 2 }).format(value);

function matchScore(requested: string | null, actual: string) {
  if (!requested) return 0;
  const a = clean(requested), b = clean(actual);
  if (a === b) return 100;
  if (a.includes(b) || b.includes(a)) return 50;
  const left = new Set(tokens(a).map(singular));
  return tokens(b).map(singular).reduce((score, token) => score + (left.has(token) ? 8 : 0), 0);
}

function resolveColumn(table: DataTable, requested: string | null) {
  return table.columns.map(column => ({ column, score: matchScore(requested, column) }))
    .sort((a, b) => b.score - a.score)[0];
}

function filterRows(table: DataTable, plan: QueryPlan) {
  let rows = table.rows;
  const applied: string[] = [];
  for (const filter of plan.filters) {
    const resolved = resolveColumn(table, filter.column);
    if (!resolved || resolved.score <= 0) continue;
    const expected = clean(filter.value);
    rows = rows.filter(row => {
      const actual = clean(row[resolved.column]);
      return filter.operator === "contains" ? actual.includes(expected) : actual === expected;
    });
    applied.push(`${resolved.column} = “${filter.value}”`);
  }
  return { rows, applied };
}

function tableScore(table: DataTable, plan: QueryPlan) {
  let score = resolveColumn(table, plan.target_column)?.score ?? 0;
  score += resolveColumn(table, plan.group_by)?.score ?? 0;
  for (const filter of plan.filters) score += resolveColumn(table, filter.column)?.score ?? 0;
  if (plan.preferred_sheet && clean(table.sheet).includes(clean(plan.preferred_sheet))) score += 80;
  if (clean(plan.subject).includes("personal") && clean(table.sheet) === "personal") score += 100;
  return score;
}

function label(table: DataTable) { return `${table.workbook} · ${table.sheet}`; }

export function executePlan(plan: QueryPlan, tables: DataTable[]): QueryResult {
  if (plan.operation === "unsupported") {
    return { answer: "La información solicitada no está disponible en los cuatro Excel. No generaré un valor sin evidencia.", interpretation: plan.interpretation, sources: [] };
  }

  const candidates = tables.map(table => ({ table, score: tableScore(table, plan) }))
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score || b.table.rows.length - a.table.rows.length);
  if (!candidates.length) {
    return { answer: "La IA interpretó la pregunta, pero no encontró columnas compatibles para ejecutar el cálculo.", interpretation: plan.interpretation, sources: [] };
  }

  if (plan.operation === "count_unique" || plan.operation === "count_rows") {
    const results = candidates.slice(0, 6).map(({ table }) => {
      const { rows, applied } = filterRows(table, plan);
      const resolved = resolveColumn(table, plan.target_column);
      const count = plan.operation === "count_unique" && resolved?.score > 0
        ? new Set(rows.map(row => String(row[resolved.column] ?? "").trim()).filter(Boolean)).size
        : rows.length;
      return { table, count, applied };
    });
    const counts = [...new Set(results.map(result => result.count))];
    if (counts.length > 1) {
      return {
        answer: "Los archivos presentan cantidades diferentes. Para evitar una respuesta incorrecta, muestro el resultado de cada fuente.",
        interpretation: plan.interpretation,
        sources: results.map(result => label(result.table)),
        table: results.map(result => ({ Archivo: result.table.workbook, Hoja: result.table.sheet, Cantidad: result.count })),
        warning: "Se detectó una discrepancia entre archivos."
      };
    }
    const filterText = results[0].applied.length ? ` con ${results[0].applied.join(" y ")}` : "";
    return {
      answer: `El resultado es ${counts[0] ?? 0} ${plan.subject}${filterText}.`,
      interpretation: plan.interpretation,
      sources: results.map(result => label(result.table))
    };
  }

  const selected = candidates[0].table;
  const { rows, applied } = filterRows(selected, plan);
  const target = resolveColumn(selected, plan.target_column);
  const group = resolveColumn(selected, plan.group_by);
  const sources = [label(selected)];

  if (plan.operation === "group_count" && group?.score > 0) {
    const id = target?.score > 0 ? target.column : null;
    const buckets = new Map<string, Set<string>>();
    rows.forEach((row, index) => {
      const key = String(row[group.column] ?? "Sin dato") || "Sin dato";
      if (!buckets.has(key)) buckets.set(key, new Set());
      buckets.get(key)!.add(id ? String(row[id] ?? index) : String(index));
    });
    const table = [...buckets].map(([key, values]) => ({ [group.column]: key, Cantidad: values.size }))
      .sort((a, b) => Number(b.Cantidad) - Number(a.Cantidad));
    return { answer: `Agrupé ${rows.length} registros por ${group.column}.`, interpretation: plan.interpretation, sources, table };
  }

  if (["sum", "average", "min", "max"].includes(plan.operation) && target?.score > 0) {
    const values = rows.map(row => Number(String(row[target.column] ?? "").replace(",", "."))).filter(Number.isFinite);
    if (!values.length) return { answer: `La columna ${target.column} no contiene valores numéricos válidos.`, interpretation: plan.interpretation, sources };
    const result = plan.operation === "sum" ? values.reduce((a, b) => a + b, 0)
      : plan.operation === "average" ? values.reduce((a, b) => a + b, 0) / values.length
      : plan.operation === "min" ? Math.min(...values) : Math.max(...values);
    const names = { sum: "total", average: "promedio", min: "mínimo", max: "máximo" } as const;
    return { answer: `El ${names[plan.operation as keyof typeof names]} de ${target.column} es ${fmt(result)}${applied.length ? `, filtrado por ${applied.join(" y ")}` : ""}.`, interpretation: plan.interpretation, sources };
  }

  const columns = selected.columns.slice(0, 10);
  const table = rows.slice(0, 20).map(row => Object.fromEntries(columns.map(column => [column, row[column]])) as Row);
  return {
    answer: rows.length ? `Encontré ${rows.length} registros relacionados. Muestro los primeros ${table.length}.` : "No encontré registros que cumplan los filtros interpretados.",
    interpretation: plan.interpretation,
    sources,
    table
  };
}
