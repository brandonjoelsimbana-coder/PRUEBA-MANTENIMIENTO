import type { DataTable, QueryPlan } from "./types";

const clean = (value: unknown) =>
  String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();

const includesAny = (question: string, terms: string[]) => terms.some(term => question.includes(term));

const columnAliases: Array<{ aliases: string[]; column: string }> = [
  { aliases: ["ot", "orden", "ordenes", "orden de trabajo", "ordenes de trabajo", "vehiculo", "vehiculos registrados"], column: "OT" },
  { aliases: ["horas productivas", "hora productiva", "horas trabajadas"], column: "Horas productivas" },
  { aliases: ["horas extra", "hora extra"], column: "Horas extra" },
  { aliases: ["costo total", "costos totales"], column: "Costo total" },
  { aliases: ["costo tecnico", "costos tecnicos", "mano de obra tecnica"], column: "Costo técnico" },
  { aliases: ["costo asesor", "costos asesores"], column: "Costo asesor" },
  { aliases: ["utilizacion", "capacidad utilizada", "porcentaje de utilizacion"], column: "Utilización" },
  { aliases: ["tecnico", "tecnicos", "mecanico", "mecanicos"], column: "Técnico" },
  { aliases: ["asesor", "asesores"], column: "Asesor" },
  { aliases: ["estado", "estados"], column: "Estado" },
  { aliases: ["tipo de vehiculo", "tipos de vehiculo"], column: "Tipo vehículo" },
  { aliases: ["tipo de trabajo", "tipos de trabajo", "mantenimiento"], column: "Tipo trabajo" },
  { aliases: ["fecha", "dia", "dias"], column: "Fecha" }
];

const groupAliases: Array<{ aliases: string[]; column: string }> = [
  { aliases: ["por tecnico", "segun tecnico"], column: "Técnico" },
  { aliases: ["por asesor", "segun asesor"], column: "Asesor" },
  { aliases: ["por estado", "segun estado"], column: "Estado" },
  { aliases: ["por tipo de vehiculo", "segun tipo de vehiculo"], column: "Tipo vehículo" },
  { aliases: ["por tipo de trabajo", "segun tipo de trabajo"], column: "Tipo trabajo" },
  { aliases: ["por fecha", "por dia", "diario"], column: "Fecha" },
  { aliases: ["por semana", "semanal"], column: "Semana" }
];

const staffRoles = [
  { aliases: ["tecnico", "tecnicos", "mecanico", "mecanicos"], value: "Técnico mecánico", subject: "técnicos mecánicos" },
  { aliases: ["asesor", "asesores"], value: "Asesor de servicio", subject: "asesores de servicio" },
  { aliases: ["lavador", "lavadores", "lavado", "personal de lavado"], value: "Lavador", subject: "lavadores" },
  { aliases: ["control de calidad", "personal de calidad", "inspector de calidad"], value: "Control de calidad", subject: "personas de control de calidad" }
];

function findTarget(question: string) {
  return columnAliases
    .map(item => ({ ...item, score: Math.max(0, ...item.aliases.filter(alias => question.includes(alias)).map(alias => alias.length)) }))
    .sort((a, b) => b.score - a.score)[0];
}

function findCategoricalFilter(question: string, tables: DataTable[]) {
  let best: { column: string; value: string; score: number } | undefined;
  for (const table of tables) {
    for (const column of table.columns) {
      const values = [...new Set(table.rows.map(row => String(row[column] ?? "").trim()).filter(Boolean))];
      for (const value of values) {
        const normalized = clean(value);
        if (!normalized || /^\d[\d.,%/$\s-]*$/.test(normalized)) continue;
        const singularVariant = normalized.endsWith("a") ? `${normalized.slice(0, -1)}as` : `${normalized}s`;
        if (question.includes(normalized) || question.includes(singularVariant)) {
          const score = normalized.length;
          if (!best || score > best.score) best = { column, value, score };
        }
      }
    }
  }
  return best;
}

export function interpretQuestion(rawQuestion: string, tables: DataTable[]): QueryPlan {
  const question = clean(rawQuestion);
  const wantsCount = includesAny(question, ["cuantos", "cuantas", "cantidad", "numero de", "total de"]);
  const wantsList = includesAny(question, ["quien", "quienes", "cuales", "lista", "muestra", "nombres"]);
  const role = staffRoles
    .map(item => ({ ...item, score: Math.max(0, ...item.aliases.filter(alias => question.includes(alias)).map(alias => alias.length)) }))
    .sort((a, b) => b.score - a.score)[0];

  if (role?.score && (wantsCount || wantsList)) {
    return {
      operation: wantsCount && !wantsList ? "count_unique" : "list",
      subject: role.subject,
      target_column: "Nombre",
      group_by: null,
      preferred_sheet: "Personal",
      filters: [{ column: "Cargo", operator: "equals", value: role.value }],
      interpretation: `Consulta sobre el personal con cargo “${role.value}”.`
    };
  }
  if (includesAny(question, ["personal", "personas", "empleados", "trabajan en el taller"]) && (wantsCount || wantsList)) {
    return {
      operation: wantsCount && !wantsList ? "count_unique" : "list",
      subject: "personas del taller",
      target_column: "Nombre",
      group_by: null,
      preferred_sheet: "Personal",
      filters: [],
      interpretation: "Consulta sobre todo el personal registrado en la hoja Personal."
    };
  }

  const group = groupAliases.find(item => item.aliases.some(alias => question.includes(alias)));
  const target = findTarget(question);
  const filter = findCategoricalFilter(question, tables);
  const filters = filter ? [{ column: filter.column, operator: "equals" as const, value: filter.value }] : [];
  let operation: QueryPlan["operation"] = "lookup";
  if (includesAny(question, ["promedio", "media"])) operation = "average";
  else if (includesAny(question, ["suma", "sumatoria", "costo total acumulado"])) operation = "sum";
  else if (includesAny(question, ["maximo", "mayor"])) operation = "max";
  else if (includesAny(question, ["minimo", "menor"])) operation = "min";
  else if (group && includesAny(question, ["distribucion", "cantidad", "cuantos", "cuantas", "por ", "segun"])) operation = "group_count";
  else if (wantsCount) operation = "count_unique";
  else if (wantsList) operation = "list";

  let targetColumn = target?.score ? target.column : null;
  let subject = target?.score ? target.column : "registros";
  if (includesAny(question, ["orden", "ordenes", " ot ", "ot hay", "vehiculo", "vehiculos"])) {
    targetColumn = "OT";
    subject = question.includes("vehiculo") ? "vehículos" : "órdenes de trabajo";
  }
  if (operation === "group_count" && !targetColumn) targetColumn = "OT";

  return {
    operation,
    subject,
    target_column: targetColumn,
    group_by: group?.column ?? null,
    preferred_sheet: null,
    filters,
    interpretation: `${operation} de ${subject}${group ? ` agrupado por ${group.column}` : ""}${filter ? ` filtrado por ${filter.column} = “${filter.value}”` : ""}.`
  };
}
