import * as XLSX from "xlsx";

export type Row = Record<string, unknown>;
export type SourceTable = {
  workbook: string;
  file: string;
  sheet: string;
  rows: Row[];
  columns: string[];
};
export type ChartKind = "bar" | "line" | "pie";
export type ChartPoint = { label: string; value: number };
export type Chat = {
  role: "user" | "assistant";
  text: string;
  table?: Row[];
  chart?: ChartPoint[];
  chartKind?: ChartKind;
  chartTitle?: string;
  autoChart?: boolean;
  sources?: string[];
};

export const clean = (s: string) =>
  s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();

const words = (s: string) =>
  clean(s).split(/[^a-z0-9]+/).filter(word => word.length > 1);

const stem = (word: string) => {
  if (word.length <= 4) return word;
  return word
    .replace(/(aciones|acion|idades|idad)$/u, "")
    .replace(/(mente)$/u, "")
    .replace(/(ores|oras|ador|adora)$/u, "")
    .replace(/(icos|icas|ico|ica)$/u, "")
    .replace(/(es|s)$/u, "");
};

const stems = (s: string) => words(s).map(stem);

export const num = (value: unknown) => {
  if (typeof value === "number") return value;
  let raw = String(value ?? "")
    .replace(/\s/g, "")
    .replace(/[$€£%]/g, "");
  if (!raw) return Number.NaN;
  const comma = raw.lastIndexOf(",");
  const dot = raw.lastIndexOf(".");
  if (comma >= 0 && dot >= 0) {
    raw = comma > dot
      ? raw.replace(/\./g, "").replace(",", ".")
      : raw.replace(/,/g, "");
  } else if (comma >= 0) {
    raw = /,\d{1,3}$/.test(raw) ? raw.replace(",", ".") : raw.replace(/,/g, "");
  }
  return Number(raw);
};

export const fmt = (value: number) =>
  new Intl.NumberFormat("es-EC", { maximumFractionDigits: 2 }).format(value);

const concepts = [
  {
    key: "ot",
    aliases: ["ot", "ots", "orden", "ordenes", "orden de trabajo", "ordenes de trabajo"],
    columns: ["ot"],
  },
  {
    key: "fecha",
    aliases: ["fecha", "dia", "dias", "jornada", "programacion", "fecha programacion"],
    columns: ["fecha", "fecha programacion", "dia"],
  },
  {
    key: "semana",
    aliases: ["semana", "semanal"],
    columns: ["semana"],
  },
  {
    key: "tecnico",
    aliases: ["tecnico", "tecnicos", "mecanico", "mecanicos", "operario", "operarios", "equipo mecanico"],
    columns: ["tecnico", "cod tecnico", "código técnico"],
  },
  {
    key: "asesor",
    aliases: ["asesor", "asesores", "recepcionista", "asesoria de servicio", "asesor de servicio"],
    columns: ["asesor", "cod asesor", "código asesor"],
  },
  {
    key: "estado",
    aliases: ["estado", "estados", "situacion de orden", "situacion de ot", "fase", "etapa"],
    columns: ["estado"],
  },
  {
    key: "tipo vehiculo",
    aliases: ["tipo de vehiculo", "tipo vehiculo", "clase de vehiculo", "categoria de vehiculo", "vehiculo", "vehiculos", "auto", "autos", "coche", "coches"],
    columns: ["tipo vehiculo", "tipo de vehiculo"],
  },
  {
    key: "modelo",
    aliases: ["modelo", "modelos"],
    columns: ["modelo"],
  },
  {
    key: "marca",
    aliases: ["marca", "marcas"],
    columns: ["marca"],
  },
  {
    key: "tipo trabajo",
    aliases: ["tipo de trabajo", "tipo trabajo", "servicio", "servicios", "mantenimiento", "mantenimientos", "tipo mantenimiento", "tipo de mantenimiento"],
    columns: ["tipo trabajo", "tipo mantenimiento", "tipo de mantenimiento"],
  },
  {
    key: "horas productivas",
    aliases: ["horas productivas", "hora productiva", "tiempo productivo", "productividad"],
    columns: ["horas productivas"],
  },
  {
    key: "horas requeridas",
    aliases: ["horas requeridas", "tiempo requerido", "horas necesarias", "mano de obra"],
    columns: ["horas requeridas"],
  },
  {
    key: "horas normales",
    aliases: ["horas normales", "horas ordinarias", "jornada normal"],
    columns: ["horas normales"],
  },
  {
    key: "horas extra",
    aliases: ["horas extra", "horas extras", "horas extraordinarias", "tiempo extra", "sobretiempo"],
    columns: ["horas extra", "horas extraordinarias"],
  },
  {
    key: "tempario",
    aliases: ["tempario", "tiempo estandar", "tiempo estándar"],
    columns: ["tempario h", "tempario"],
  },
  {
    key: "costo tecnico",
    aliases: ["costo tecnico", "costo del tecnico", "costo mecanico", "mano de obra tecnica"],
    columns: ["costo tecnico"],
  },
  {
    key: "costo asesor",
    aliases: ["costo asesor", "costo del asesor"],
    columns: ["costo asesor"],
  },
  {
    key: "costo lavado",
    aliases: ["costo lavado", "costo de lavado"],
    columns: ["costo lavado"],
  },
  {
    key: "costo control calidad",
    aliases: ["costo control calidad", "costo de control de calidad", "costo calidad", "costo cc"],
    columns: ["costo control calidad", "costo cc"],
  },
  {
    key: "costo total",
    aliases: ["costo total", "costos totales", "gasto total", "coste total"],
    columns: ["costo total"],
  },
  {
    key: "precio estimado",
    aliases: ["precio estimado", "precio", "ingreso estimado"],
    columns: ["precio estimado"],
  },
  {
    key: "ganancia estimada",
    aliases: ["ganancia estimada", "ganancia", "utilidad estimada", "beneficio"],
    columns: ["ganancia estimada"],
  },
  {
    key: "capacidad",
    aliases: ["capacidad", "capacidad taller", "capacidad total", "capacidad disponible", "disponibilidad", "reserva"],
    columns: ["capacidad normal", "capacidad programable", "capacidad total", "disponible", "reserva 15", "capacidad tecnico"],
  },
  {
    key: "utilizacion",
    aliases: ["utilizacion", "uso de capacidad", "ocupacion", "carga del taller"],
    columns: ["utilizacion", "utilizacion capacidad"],
  },
  {
    key: "cumplimiento",
    aliases: ["cumplimiento", "a tiempo", "retraso", "retrasos", "cumple entrega", "entrega cumplida"],
    columns: ["cumplimiento", "cumple entrega"],
  },
  {
    key: "situacion",
    aliases: ["situacion", "terminada", "terminadas", "pendiente", "pendientes"],
    columns: ["situacion"],
  },
  {
    key: "bahia",
    aliases: ["bahia", "bahias", "puesto", "puestos", "estacion de trabajo"],
    columns: ["bahia"],
  },
  {
    key: "cliente",
    aliases: ["cliente", "clientes", "propietario"],
    columns: ["cliente"],
  },
  {
    key: "placa",
    aliases: ["placa", "matricula"],
    columns: ["placa"],
  },
  {
    key: "conflicto",
    aliases: ["conflicto", "conflictos", "solapamiento", "sobrecarga", "problema"],
    columns: ["conflicto", "diagnostico", "indicador"],
  },
];

const intentWords = new Set([
  "que", "cual", "cuales", "cuanto", "cuantos", "cuanta", "cuantas", "dime",
  "mostrar", "muestra", "lista", "listar", "total", "suma", "promedio", "media",
  "maximo", "minimo", "mayor", "menor", "cantidad", "conteo", "grafico", "grafica",
  "distribucion", "por", "segun", "del", "los", "las", "una", "unos", "datos",
  "excel", "archivo", "hay", "tiene", "existe", "existen", "son", "esta", "estan",
]);

const operationAliases = {
  count: ["cuanto", "cuantos", "cuanta", "cuantas", "cantidad", "numero de", "conteo", "cuenta", "existen", "hay", "tiene el taller"],
  average: ["promedio", "media", "valor medio", "media aritmetica"],
  total: ["total", "suma", "sumatoria", "acumulado", "acumulada"],
  max: ["maximo", "maxima", "mayor", "mas", "mas alto", "mas alta", "top"],
  min: ["minimo", "minima", "menor", "menos", "mas bajo", "mas baja"],
  list: ["muestra", "mostrar", "lista", "listar", "detalle", "detalla", "cuales", "quien", "quienes", "nombres", "registros"],
  group: [" por ", " segun ", "distribucion", "desglose", "composicion", "agrupa", "agrupado"],
  chart: ["grafica", "grafico", "diagrama", "visualiza", "visualizacion", "representa", "chart"],
};

const stateAliases: Record<string, string[]> = {
  "En desarrollo": ["en desarrollo", "en proceso", "en curso", "trabajando", "activas", "activos"],
  Planificada: ["planificada", "planificadas", "programada", "programadas"],
  Proyectada: ["proyectada", "proyectadas", "proyectado", "proyectados"],
  Prevista: ["prevista", "previstas", "previsto", "previstos"],
  Parada: ["parada", "paradas", "detenida", "detenidas", "bloqueada", "bloqueadas"],
  Pendiente: ["pendiente", "pendientes"],
  Terminada: ["terminada", "terminadas", "finalizada", "finalizadas", "completada", "completadas"],
  Retraso: ["retraso", "retrasadas", "tarde", "incumplidas"],
  "A tiempo": ["a tiempo", "puntuales", "cumplidas"],
  CONFLICTO: ["conflicto", "conflictos", "solapada", "solapadas"],
  OK: ["sin conflicto", "correctas", "validas"],
  SOBRECARGA: ["sobrecarga", "sobrecargada", "sobrecargadas"],
};

export function expandQuestion(question: string) {
  const base = ` ${clean(question)} `;
  const additions = new Set<string>();
  for (const concept of concepts) {
    if (concept.aliases.some(alias => base.includes(` ${clean(alias)} `) || base.includes(clean(alias)))) {
      additions.add(concept.key);
    }
  }
  Object.entries(operationAliases).forEach(([operation, aliases]) => {
    if (aliases.some(alias => base.includes(clean(alias)))) additions.add(operation);
  });
  return `${base.trim()} ${[...additions].join(" ")}`.trim();
}

function parseSheet(ws: XLSX.WorkSheet, forceFirstHeader = false): Row[] {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    defval: "",
    raw: false,
  });
  if (!matrix.length) return [];
  const limit = Math.min(15, matrix.length);
  let headerIndex = 0;
  let best = -1;
  for (let index = 0; index < limit; index++) {
    const values = (matrix[index] ?? []).map(value => String(value ?? "").trim()).filter(Boolean);
    const score = values.length * 3 + new Set(values.map(clean)).size - (values.length <= 1 ? 8 : 0);
    if (score > best) {
      best = score;
      headerIndex = index;
    }
  }
  if (forceFirstHeader) headerIndex = 0;
  const raw = (matrix[headerIndex] ?? []).map((value, index) =>
    String(value ?? "").trim() || `Columna ${index + 1}`);
  const seen = new Map<string, number>();
  const headers = raw.map(header => {
    const count = (seen.get(header) ?? 0) + 1;
    seen.set(header, count);
    return count === 1 ? header : `${header} ${count}`;
  });
  return matrix.slice(headerIndex + 1)
    .filter(row => row.some(value => String(value ?? "").trim() !== ""))
    .map(row => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
}

export function parseWorkbook(data: ArrayBuffer, workbook: string, file: string): SourceTable[] {
  const book = XLSX.read(data, { type: "array", cellDates: true });
  return book.SheetNames.map(sheet => {
    const rows = parseSheet(book.Sheets[sheet], clean(sheet) === "resumenes");
    return { workbook, file, sheet, rows, columns: Object.keys(rows[0] ?? {}) };
  }).filter(source => source.rows.length && source.columns.length);
}

export function sourceLabel(source: SourceTable) {
  return `${source.workbook} · ${source.sheet}`;
}

function conceptForColumn(column: string) {
  const normalized = clean(column).replace(/[^a-z0-9]+/g, " ").trim();
  return concepts.find(concept =>
    concept.columns.some(candidate => {
      const normalizedCandidate = clean(candidate);
      return normalized === normalizedCandidate || normalized.includes(normalizedCandidate);
    }));
}

function columnMatch(question: string, column: string) {
  const query = expandQuestion(question);
  const normalizedColumn = clean(column);
  let score = query.includes(normalizedColumn) ? 14 + normalizedColumn.length / 10 : 0;
  const queryStems = new Set(stems(query));
  score += stems(normalizedColumn).reduce((total, word) => total + (queryStems.has(word) ? 4 : 0), 0);
  const concept = conceptForColumn(column);
  if (concept && (query.includes(concept.key) || concept.aliases.some(alias => query.includes(clean(alias))))) {
    score += 15;
  }
  return score;
}

function sourcePriority(source: SourceTable) {
  const sheet = clean(source.sheet);
  if (sheet === "personal" || sheet === "diccionario" || sheet === "parametros") return 8;
  if (sheet === "base_datos" || sheet === "ordenes_trabajo" || sheet === "planning_diario") return 7;
  if (sheet === "capacidad_taller" || sheet === "conflictos") return 6;
  if (sheet === "resumen_base" || sheet === "resumenes" || sheet === "dashboard") return 4;
  if (sheet === "portada") return -4;
  return 0;
}

function sourceScore(question: string, source: SourceTable) {
  const query = expandQuestion(question);
  let score = source.columns.reduce((total, column) => total + columnMatch(query, column), 0);
  score += words(`${source.workbook} ${source.sheet}`).reduce(
    (total, word) => total + (query.includes(word) ? 2 : 0), 0);
  const queryTerms = words(query).filter(word => !intentWords.has(word));
  const sample = source.rows.slice(0, 260);
  for (const term of queryTerms) {
    if (sample.some(row =>
      Object.values(row).some(value => clean(String(value ?? "")).includes(term)))) {
      score += 4;
    }
  }
  return score > 0 ? score + sourcePriority(source) : 0;
}

function bestColumn(question: string, sources: SourceTable[], numericOnly = false) {
  let best: { source: SourceTable; column: string; score: number } | undefined;
  for (const source of sources) {
    for (const column of source.columns) {
      if (numericOnly && !source.rows.some(row => Number.isFinite(num(row[column])))) continue;
      const semanticScore = columnMatch(question, column);
      if (semanticScore <= 0) continue;
      const score = semanticScore + sourcePriority(source) / 10;
      if (
        score > 0 &&
        (!best || score > best.score ||
          (score === best.score && source.rows.length > best.source.rows.length))
      ) {
        best = { source, column, score };
      }
    }
  }
  return best;
}

type AppliedFilter = { column: string; value: string };

function filtersFor(question: string, source: SourceTable): AppliedFilter[] {
  const query = ` ${clean(question)} `;
  const queryStems = new Set(stems(query));
  const matches: { column: string; value: string; score: number }[] = [];
  for (const column of source.columns) {
    const values = [...new Set(
      source.rows
        .map(row => String(row[column] ?? "").trim())
        .filter(value => value && value.length < 90),
    )].slice(0, 500);
    for (const value of values) {
      const normalizedValue = clean(value);
      const valueWords = words(normalizedValue);
      const valueStems = valueWords.map(stem);
      const numericLike = /^[$€£]?\s*-?\d[\d.,%/\s:-]*$/.test(value);
      const literal = normalizedValue.length >= 3 && query.includes(normalizedValue);
      const semantic = !numericLike && valueWords.length > 0 &&
        valueStems.every(valueStem => queryStems.has(valueStem));
      const numericWithColumn = numericLike &&
        new RegExp(`(^|[^0-9])${normalizedValue.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}([^0-9]|$)`).test(query) &&
        columnMatch(query, column) > 5;
      if (literal || semantic || numericWithColumn) {
        matches.push({ column, value, score: normalizedValue.length + columnMatch(query, column) });
      }
    }
  }

  for (const [canonical, aliases] of Object.entries(stateAliases)) {
    if (!aliases.some(alias => query.includes(clean(alias)))) continue;
    for (const column of source.columns) {
      const exact = source.rows.find(row => clean(String(row[column] ?? "")) === clean(canonical));
      if (exact) matches.push({ column, value: canonical, score: 100 + columnMatch(query, column) });
    }
  }

  const byColumn = new Map<string, { column: string; value: string; score: number }>();
  for (const match of matches) {
    const current = byColumn.get(match.column);
    if (!current || match.score > current.score) byColumn.set(match.column, match);
  }
  return [...byColumn.values()]
    .sort((a, b) => b.score - a.score)
    .slice(0, 3)
    .map(({ column, value }) => ({ column, value }));
}

function applyFilters(rows: Row[], filters: AppliedFilter[]) {
  return filters.length
    ? rows.filter(row => filters.every(filter =>
      clean(String(row[filter.column] ?? "")) === clean(filter.value)))
    : rows;
}

function filterText(filters: AppliedFilter[]) {
  return filters.length
    ? ` con ${filters.map(filter => `${filter.column} = “${filter.value}”`).join(" y ")}`
    : "";
}

function operationRequested(query: string, operation: keyof typeof operationAliases) {
  const normalized = ` ${clean(query)} `;
  return operationAliases[operation].some(alias => normalized.includes(clean(alias)));
}

function wantsHelp(query: string) {
  const q = clean(query);
  return [
    "que puedo preguntar", "que preguntas", "preguntas disponibles", "ejemplos de preguntas",
    "ayuda", "como preguntar", "consultas disponibles", "que sabes",
  ].some(phrase => q.includes(phrase));
}

function helpCatalog(sources: SourceTable[]): Row[] {
  const columns = new Set(sources.flatMap(source => source.columns.map(clean)));
  const groups: { topic: string; required: string[]; examples: string[] }[] = [
    {
      topic: "Órdenes y estados",
      required: ["ot", "estado"],
      examples: [
        "¿Cuántas órdenes de trabajo hay?",
        "¿Cuántas OT están paradas?",
        "Gráfica de órdenes por estado",
      ],
    },
    {
      topic: "Personal",
      required: ["nombre", "cargo"],
      examples: [
        "¿Cuántos técnicos, asesores, lavadores y personas de control de calidad hay?",
        "¿Quiénes son los técnicos?",
        "Lista todo el personal por cargo",
      ],
    },
    {
      topic: "Vehículos y servicios",
      required: ["tipo vehiculo"],
      examples: [
        "¿Cuántos vehículos hay por tipo?",
        "¿Qué modelos están registrados?",
        "Gráfica de mantenimientos por tipo de trabajo",
      ],
    },
    {
      topic: "Horas y productividad",
      required: ["horas productivas"],
      examples: [
        "Promedio de horas productivas por técnico",
        "Total de horas extra por fecha",
        "¿Qué técnico acumula más horas productivas?",
      ],
    },
    {
      topic: "Costos",
      required: ["costo total"],
      examples: [
        "¿Cuál es el costo total de todas las OT?",
        "Costo promedio por tipo de vehículo",
        "Gráfica del costo técnico por técnico",
      ],
    },
    {
      topic: "Capacidad",
      required: ["utilizacion"],
      examples: [
        "Promedio de utilización del taller",
        "¿Qué día tuvo mayor carga asignada?",
        "Gráfica de capacidad total por fecha",
      ],
    },
    {
      topic: "Cumplimiento y conflictos",
      required: ["cumplimiento"],
      examples: [
        "¿Cuántas órdenes tuvieron retraso?",
        "Distribución de cumplimiento",
        "¿Cuántas órdenes tienen horas extra?",
      ],
    },
    {
      topic: "Definiciones y parámetros",
      required: ["campo"],
      examples: [
        "¿Qué significa costo técnico?",
        "¿Cómo se calcula la ganancia estimada?",
        "Muestra los parámetros del taller",
      ],
    },
  ];
  return groups
    .filter(group => group.required.some(required =>
      [...columns].some(column => column.includes(required))))
    .map(group => ({
      Tema: group.topic,
      "Variantes admitidas": group.examples.join(" · "),
    }));
}

function findGroupColumn(
  question: string,
  source: SourceTable,
  target?: string,
  allowInference = true,
) {
  const q = clean(question);
  const explicitGroup = q.match(/(?:\bpor\b|\bsegun\b)\s+(.+)$/)?.[1];
  const groupQuery = explicitGroup ?? q;
  let candidates = source.columns
    .filter(column => column !== target)
    .filter(column => source.rows.some(row => {
      const value = String(row[column] ?? "").trim();
      return value && !Number.isFinite(num(value));
    }))
    .map(column => ({ column, score: columnMatch(groupQuery, column) }))
    .sort((a, b) => b.score - a.score);
  if (explicitGroup && candidates[0]?.score > 0) return candidates[0].column;
  if (!allowInference) return undefined;
  candidates = candidates
    .map(candidate => ({
      ...candidate,
      score: candidate.score + (conceptForColumn(candidate.column) &&
        expandQuestion(q).includes(conceptForColumn(candidate.column)!.key) ? 10 : 0),
    }))
    .sort((a, b) => b.score - a.score);
  return candidates[0]?.score > 10 ? candidates[0].column : undefined;
}

function personnelAnswer(question: string, sources: SourceTable[], wantsCount: boolean, wantsList: boolean) {
  const q = clean(question);
  const personal = sources.find(source =>
    clean(source.sheet) === "personal" &&
    source.columns.some(column => clean(column) === "nombre") &&
    source.columns.some(column => clean(column) === "cargo"));
  if (!personal) return undefined;
  const nameColumn = personal.columns.find(column => clean(column) === "nombre")!;
  const roleColumn = personal.columns.find(column => clean(column) === "cargo")!;
  const codeColumn = personal.columns.find(column => clean(column) === "codigo");
  const roles = [
    {
      aliases: ["tecnico", "tecnicos", "mecanico", "mecanicos", "personal tecnico", "equipo mecanico", "operario", "operarios"],
      cargo: "tecnico mecanico",
      singular: "técnico mecánico",
      plural: "técnicos mecánicos",
    },
    {
      aliases: ["asesor", "asesores", "asesor de servicio", "recepcionista", "recepcionistas"],
      cargo: "asesor de servicio",
      singular: "asesor de servicio",
      plural: "asesores de servicio",
    },
    {
      aliases: ["lavador", "lavadores", "lavado", "personal de lavado", "quien lava", "encargado de lavado"],
      cargo: "lavador",
      singular: "lavador",
      plural: "lavadores",
    },
    {
      aliases: ["control de calidad", "calidad", "inspector de calidad", "inspectores de calidad", "personal de control de calidad", "quien revisa", "encargado de calidad"],
      cargo: "control de calidad",
      singular: "persona de control de calidad",
      plural: "personas de control de calidad",
    },
  ];
  const role = roles
    .map(item => ({
      ...item,
      match: Math.max(...item.aliases.filter(alias => q.includes(alias)).map(alias => alias.length), 0),
    }))
    .sort((a, b) => b.match - a.match)[0];
  const asksPeople = wantsCount || wantsList ||
    ["personal", "personas", "empleados", "equipo del taller", "trabajan en el taller"].some(term => q.includes(term));
  if (!asksPeople) return undefined;
  if (role?.match) {
    const people = personal.rows
      .filter(row => clean(String(row[roleColumn] ?? "")).includes(role.cargo))
      .map(row => ({
        ...(codeColumn ? { Código: String(row[codeColumn] ?? "") } : {}),
        Nombre: String(row[nameColumn] ?? ""),
        Cargo: String(row[roleColumn] ?? ""),
      }))
      .filter(person => person.Nombre);
    const label = people.length === 1 ? role.singular : role.plural;
    return {
      role: "assistant" as const,
      text: `Hay ${people.length} ${label}${people.length ? `: ${people.map(person => person.Nombre).join(", ")}` : ""}.`,
      table: people,
      chart: [{ label: role.plural, value: people.length }],
      chartTitle: `Cantidad de ${role.plural}`,
      sources: [sourceLabel(personal)],
    };
  }
  const asksAll = ["personal", "personas", "empleados", "equipo del taller", "trabajan en el taller"].some(term => q.includes(term));
  if (!asksAll) return undefined;
  const people = personal.rows
    .map(row => ({
      Nombre: String(row[nameColumn] ?? "").trim(),
      Cargo: String(row[roleColumn] ?? "").trim(),
    }))
    .filter(person => person.Nombre && person.Cargo);
  const roleCounts = new Map<string, number>();
  people.forEach(person => roleCounts.set(person.Cargo, (roleCounts.get(person.Cargo) ?? 0) + 1));
  return {
    role: "assistant" as const,
    text: `El taller tiene ${people.length} personas registradas. El detalle está organizado por cargo.`,
    table: people,
    chart: [...roleCounts].map(([label, value]) => ({ label, value })),
    chartTitle: "Personal por cargo",
    sources: [sourceLabel(personal)],
  };
}

function definitionAnswer(question: string, sources: SourceTable[]) {
  const q = clean(question);
  if (!/\b(que es|que significa|define|definicion|como se calcula|explica)\b/.test(q)) {
    return undefined;
  }
  const dictionaries = sources.filter(source =>
    source.columns.some(column => clean(column) === "campo") &&
    source.columns.some(column => clean(column) === "descripcion"));
  for (const dictionary of dictionaries) {
    const fieldColumn = dictionary.columns.find(column => clean(column) === "campo")!;
    const descriptionColumn = dictionary.columns.find(column => clean(column) === "descripcion")!;
    const candidates = dictionary.rows
      .map(row => ({
        field: String(row[fieldColumn] ?? "").trim(),
        description: String(row[descriptionColumn] ?? "").trim(),
      }))
      .filter(item => item.field && item.description)
      .map(item => ({
        ...item,
        score: columnMatch(expandQuestion(question), item.field) +
          (q.includes(clean(item.field)) ? 30 : 0),
      }))
      .sort((a, b) => b.score - a.score);
    if (candidates[0]?.score > 4) {
      const item = candidates[0];
      return {
        role: "assistant" as const,
        text: `${item.field}: ${item.description}`,
        table: [{ Campo: item.field, Descripción: item.description }],
        sources: [sourceLabel(dictionary)],
      };
    }
  }
  return undefined;
}

function operationalCount(question: string, sources: SourceTable[], wantsCount: boolean) {
  const q = clean(question);
  const asksOrders = /\bot\b/.test(q) || q.includes("orden de trabajo") || q.includes("ordenes");
  const asksVehicles = q.includes("vehiculo") || q.includes("autos") || q.includes("coches");
  if (!wantsCount || (!asksOrders && !asksVehicles)) return undefined;
  const operational = sources.filter(source =>
    source.columns.some(column => clean(column) === "ot") &&
    ["base_datos", "ordenes_trabajo", "planning_diario", "planning_ia"].includes(clean(source.sheet)));
  if (!operational.length) return undefined;
  const rowsBySource = operational.map(source => {
    const id = source.columns.find(column => clean(column) === "ot")!;
    const filters = filtersFor(question, source).filter(filter => clean(filter.column) !== "ot");
    const filtered = applyFilters(source.rows, filters);
    return {
      source,
      filters,
      count: new Set(filtered.map(row => String(row[id] ?? "").trim()).filter(Boolean)).size,
    };
  });
  const sourcesWithFilters = rowsBySource.filter(item => item.filters.length);
  const comparable = (sourcesWithFilters.length ? sourcesWithFilters : rowsBySource)
    .filter(item => item.count > 0 || !item.filters.length);
  const counts = [...new Set(comparable.map(item => item.count))];
  const entity = asksVehicles && !asksOrders ? "vehículos" : "órdenes de trabajo";
  const allFilters = new Map<string, string>();
  comparable.forEach(item => item.filters.forEach(filter => allFilters.set(filter.column, filter.value)));
  const filters = [...allFilters].map(([column, value]) => ({ column, value }));
  if (counts.length > 1) {
    return {
      role: "assistant" as const,
      text: `Las fuentes no coinciden en la cantidad de ${entity}${filterText(filters)}. Muestro cada valor para evitar presentar una cifra dudosa.`,
      table: comparable.map(item => ({
        Archivo: item.source.workbook,
        Hoja: item.source.sheet,
        Cantidad: item.count,
      })),
      sources: comparable.map(item => sourceLabel(item.source)),
    };
  }
  const count = counts[0] ?? 0;
  return {
    role: "assistant" as const,
    text: `El taller tiene ${count} ${entity}${filterText(filters)}. Se contaron identificadores OT únicos y no se duplicaron registros entre libros.`,
    table: [{ Indicador: entity, Cantidad: count }],
    sources: comparable.map(item => sourceLabel(item.source)),
  };
}

function aggregate(
  rows: Row[],
  target: string,
  operation: "average" | "total" | "max" | "min",
) {
  const values = rows.map(row => num(row[target])).filter(Number.isFinite);
  if (!values.length) return undefined;
  if (operation === "average") return values.reduce((a, b) => a + b, 0) / values.length;
  if (operation === "total") return values.reduce((a, b) => a + b, 0);
  if (operation === "min") return Math.min(...values);
  return Math.max(...values);
}

function operationLabel(operation: "average" | "total" | "max" | "min") {
  return operation === "average" ? "Promedio"
    : operation === "total" ? "Total"
      : operation === "min" ? "Mínimo"
        : "Máximo";
}

export function answerAcross(question: string, sources: SourceTable[]): Chat {
  const q = clean(question);
  const expanded = expandQuestion(question);
  const coverage = "Los archivos contienen órdenes, fechas, estados, personal, vehículos, trabajos, horas, capacidad, costos, cumplimiento, conflictos, temparios y parámetros.";
  if (!sources.length) {
    return { role: "assistant", text: "Todavía estoy cargando las bases de datos. Intenta nuevamente en unos segundos." };
  }

  const wantsCount = operationRequested(q, "count");
  const wantsAverage = operationRequested(q, "average");
  const wantsTotal = operationRequested(q, "total");
  const wantsMax = operationRequested(q, "max");
  const wantsMin = operationRequested(q, "min");
  const wantsList = operationRequested(q, "list");
  const wantsGroup = operationRequested(` ${q} `, "group");
  const wantsChart = operationRequested(q, "chart");
  const asksOrders = /\bot\b/.test(q) || q.includes("orden");

  if (wantsHelp(q)) {
    return {
      role: "assistant",
      text: "Puedes preguntar con lenguaje natural. Reconozco singular, plural, sinónimos, abreviaturas como OT y órdenes como contar, listar, filtrar, sumar, promediar, comparar o graficar.",
      table: helpCatalog(sources),
      sources: [...new Set(sources.map(source => source.workbook))],
    };
  }

  if (
    q.includes("que informacion") || q.includes("que datos") || q.includes("contenido") ||
    q.includes("estructura") || q.includes("columnas") || q.includes("campos")
  ) {
    return {
      role: "assistant",
      text: `Analicé ${sources.length} hojas de los cuatro libros. ${coverage}`,
      table: sources.map(source => ({
        Archivo: source.workbook,
        Hoja: source.sheet,
        Registros: source.rows.length,
        Campos: source.columns.slice(0, 10).join(", "),
      })),
      sources: [...new Set(sources.map(source => source.workbook))],
    };
  }

  const definition = definitionAnswer(question, sources);
  if (definition) return definition;

  const personnel = personnelAnswer(question, sources, wantsCount, wantsList);
  if (personnel) return personnel;

  const directCount = operationalCount(question, sources, wantsCount);
  if (directCount && !wantsGroup) return directCount;

  const ranked = sources
    .map(source => ({ source, score: sourceScore(expanded, source) }))
    .sort((a, b) => b.score - a.score);
  const relevant = ranked.filter(item => item.score > 3).map(item => item.source);
  const pool = relevant.length ? relevant : sources;
  const numericConceptKeys = [
    "horas productivas", "horas requeridas", "horas normales", "horas extra",
    "tempario", "costo tecnico", "costo asesor", "costo lavado",
    "costo control calidad", "costo total", "precio estimado", "ganancia estimada",
    "capacidad", "utilizacion",
  ];
  const hasNumericConcept = numericConceptKeys.some(key => expanded.includes(key));
  const needsNumericTarget =
    wantsAverage || wantsTotal ||
    ((wantsMax || wantsMin) && (hasNumericConcept || !asksOrders)) ||
    (wantsChart && hasNumericConcept);
  const numericTarget = needsNumericTarget ? bestColumn(expanded, pool, true) : undefined;
  const anyTarget = bestColumn(expanded, pool, false);

  if (!relevant.length && !anyTarget) {
    return {
      role: "assistant",
      text: `No encontré información suficiente en los cuatro Excel para responder sin inventar datos. ${coverage} Prueba con “¿Qué puedo preguntar?”.`,
    };
  }

  let targetSource = numericTarget?.source ?? anyTarget?.source ?? relevant[0];
  if (wantsCount || wantsList || wantsChart) {
    const detailed = relevant
      .filter(source => ["personal", "base_datos", "ordenes_trabajo", "planning_diario", "planning_ia", "capacidad_taller", "conflictos", "parametros", "temparios", "vehiculos"].includes(clean(source.sheet)))
      .sort((a, b) => sourceScore(expanded, b) - sourceScore(expanded, a));
    if (detailed[0]) targetSource = detailed[0];
  }
  if (!targetSource) {
    return { role: "assistant", text: `No pude relacionar la pregunta con un campo disponible. ${coverage}` };
  }

  const filters = filtersFor(question, targetSource);
  const filteredRows = applyFilters(targetSource.rows, filters);
  const target = numericTarget?.source === targetSource ? numericTarget.column : undefined;
  const group = findGroupColumn(
    question,
    targetSource,
    target,
    wantsGroup || wantsChart || wantsMax || wantsMin,
  );
  const source = [sourceLabel(targetSource)];

  const defaultGroupedCount = group && (wantsCount || (wantsChart && (!target || asksOrders)) || ((wantsMax || wantsMin) && asksOrders));
  if (defaultGroupedCount) {
    const groupConcept = conceptForColumn(group)?.key;
    const comparableSources = sources
      .filter(candidate => candidate.columns.some(column => clean(column) === "ot"))
      .map(candidate => ({
        source: candidate,
        groupColumn: candidate.columns.find(column =>
          clean(column) === clean(group) ||
          (groupConcept && conceptForColumn(column)?.key === groupConcept)),
      }))
      .filter((candidate): candidate is { source: SourceTable; groupColumn: string } =>
        Boolean(candidate.groupColumn))
      .filter(candidate =>
        ["base_datos", "ordenes_trabajo", "planning_diario", "planning_ia"].includes(clean(candidate.source.sheet)));
    const distributions = comparableSources.map(candidate => {
      const idColumn = candidate.source.columns.find(column => clean(column) === "ot")!;
      const candidateFilters = filtersFor(question, candidate.source)
        .filter(filter => clean(filter.column) !== clean(candidate.groupColumn));
      const rows = applyFilters(candidate.source.rows, candidateFilters);
      const buckets = new Map<string, Set<string>>();
      rows.forEach((row, index) => {
        const label = String(row[candidate.groupColumn] ?? "Sin dato") || "Sin dato";
        const id = String(row[idColumn] ?? index);
        if (!buckets.has(label)) buckets.set(label, new Set());
        buckets.get(label)!.add(id);
      });
      return {
        ...candidate,
        data: [...buckets]
          .map(([label, ids]) => ({ label, value: ids.size }))
          .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label)),
      };
    });
    const signatures = [...new Set(distributions.map(distribution =>
      JSON.stringify([...distribution.data].sort((a, b) => a.label.localeCompare(b.label)))))];
    if (distributions.length > 1 && signatures.length > 1 && !((wantsMax || wantsMin) && !wantsGroup)) {
      const table = distributions.flatMap(distribution =>
        distribution.data.map(item => ({
          Archivo: distribution.source.workbook,
          Hoja: distribution.source.sheet,
          [group]: item.label,
          Cantidad: item.value,
        })));
      return {
        role: "assistant",
        text: `Las fuentes no coinciden en la distribución de órdenes por ${group}. Muestro el desglose de cada libro para que la diferencia sea visible y no presentar una sola cifra como definitiva.`,
        table,
        chart: table.slice(0, 18).map(row => ({
          label: `${String(row.Archivo).replace(/ · .*/, "")}: ${String(row[group])}`,
          value: Number(row.Cantidad),
        })),
        chartTitle: `Comparación de órdenes por ${group} y fuente`,
        autoChart: wantsChart,
        sources: distributions.map(distribution => sourceLabel(distribution.source)),
      };
    }
    const chosen = distributions.find(distribution => distribution.source === targetSource) ?? distributions[0];
    let data = chosen?.data ?? [];
    if ((wantsMax || wantsMin) && !wantsGroup) {
      data = [wantsMin ? data[data.length - 1] : data[0]].filter(Boolean);
    }
    return {
      role: "assistant",
      text: (wantsMax || wantsMin) && !wantsGroup
        ? `${data[0]?.label ?? "Sin dato"} es ${wantsMin ? "el grupo con menos" : "el grupo con más"} órdenes: ${data[0]?.value ?? 0}.`
        : `Encontré ${filteredRows.length} registros${filterText(filters)} y los agrupé por ${group}.${distributions.length > 1 ? ` Las ${distributions.length} fuentes operativas coinciden.` : ""}`,
      table: data.map(item => ({ [group]: item.label, Cantidad: item.value })),
      sources: distributions.length
        ? distributions.map(distribution => sourceLabel(distribution.source))
        : source,
    };
  }

  if (wantsCount) {
    const idColumn = targetSource.columns.find(column => clean(column) === "ot");
    const count = idColumn
      ? new Set(filteredRows.map(row => String(row[idColumn] ?? "")).filter(Boolean)).size
      : filteredRows.length;
    return {
      role: "assistant",
      text: `Encontré ${count} ${idColumn ? "órdenes únicas" : "registros"}${filterText(filters)}.`,
      table: [{ Indicador: anyTarget?.column ?? targetSource.sheet, Cantidad: count }],
      sources: source,
    };
  }

  const operation: "average" | "total" | "max" | "min" | undefined =
    wantsAverage ? "average"
      : wantsTotal ? "total"
        : wantsMin ? "min"
          : wantsMax ? "max"
            : wantsChart && target ? "total"
              : undefined;

  if (operation && !target) {
    return {
      role: "assistant",
      text: `La pregunta solicita un cálculo, pero no identifica una columna numérica. Prueba con “promedio de Horas productivas”, “total de Costo técnico” o “máximo de Utilización”.`,
      sources: source,
    };
  }

  if (target && operation) {
    if (group) {
      const buckets = new Map<string, Row[]>();
      filteredRows.forEach(row => {
        const label = String(row[group] ?? "Sin dato") || "Sin dato";
        buckets.set(label, [...(buckets.get(label) ?? []), row]);
      });
      const inferredRanking = (wantsMax || wantsMin) && !wantsGroup;
      const perGroupOperation: "average" | "total" | "max" | "min" =
        inferredRanking ? "total" : operation;
      let data = [...buckets]
        .map(([label, rows]) => ({
          label,
          value: aggregate(rows, target, perGroupOperation),
        }))
        .filter((item): item is { label: string; value: number } => item.value !== undefined)
        .sort((a, b) => b.value - a.value);
      if (inferredRanking) data = [wantsMin ? data[data.length - 1] : data[0]].filter(Boolean);
      const label = inferredRanking ? "Total" : operationLabel(perGroupOperation);
      return {
        role: "assistant",
        text: inferredRanking
          ? `${data[0]?.label ?? "Sin dato"} tiene el ${wantsMin ? "menor" : "mayor"} acumulado de ${target}: ${fmt(data[0]?.value ?? 0)}.`
          : `Calculé ${label.toLowerCase()} de ${target} por ${group} usando ${data.length} grupos.`,
        table: data.map(item => ({
          [group]: item.label,
          [`${label} de ${target}`]: Number(item.value.toFixed(2)),
        })),
        sources: source,
      };
    }
    const targetConcept = conceptForColumn(target)?.key;
    const targetScore = columnMatch(question, target);
    const comparable = sources
      .filter(candidate =>
        ["base_datos", "ordenes_trabajo", "planning_diario", "planning_ia", "capacidad_taller"].includes(clean(candidate.sheet)))
      .map(candidate => {
        const candidateColumn = candidate.columns
          .filter(column => candidate.rows.some(row => Number.isFinite(num(row[column]))))
          .filter(column =>
            clean(column) === clean(target) ||
            (targetConcept && conceptForColumn(column)?.key === targetConcept))
          .map(column => ({ column, score: columnMatch(question, column) }))
          .sort((a, b) => b.score - a.score)[0];
        if (!candidateColumn || candidateColumn.score < targetScore - 3) return undefined;
        const candidateFilters = filtersFor(question, candidate);
        const rows = applyFilters(candidate.rows, candidateFilters);
        const result = aggregate(rows, candidateColumn.column, operation);
        return result === undefined
          ? undefined
          : { source: candidate, column: candidateColumn.column, result };
      })
      .filter((item): item is { source: SourceTable; column: string; result: number } => Boolean(item));
    if (comparable.length > 1) {
      const distinctResults = [...new Set(comparable.map(item => item.result.toFixed(4)))];
      if (distinctResults.length > 1) {
        const label = operationLabel(operation);
        const table = comparable.map(item => ({
          Archivo: item.source.workbook,
          Hoja: item.source.sheet,
          Campo: item.column,
          [`${label}`]: Number(item.result.toFixed(2)),
        }));
        return {
          role: "assistant",
          text: `Los archivos no coinciden en el ${label.toLowerCase()} de ${target}. Muestro el cálculo de cada fuente para no mezclar definiciones ni presentar un valor dudoso.`,
          table,
          chart: table.map(row => ({
            label: `${String(row.Archivo).replace(/ · .*/, "")}: ${String(row.Hoja)}`,
            value: Number(row[label]),
          })),
          chartTitle: `${label} de ${target} por fuente`,
          autoChart: wantsChart,
          sources: comparable.map(item => sourceLabel(item.source)),
        };
      }
      const commonResult = comparable[0].result;
      return {
        role: "assistant",
        text: `El ${operationLabel(operation).toLowerCase()} de ${target} es ${fmt(commonResult)}. Las ${comparable.length} fuentes con ese campo coinciden.`,
        table: [{ Indicador: `${operationLabel(operation)} de ${target}`, Valor: Number(commonResult.toFixed(2)) }],
        sources: comparable.map(item => sourceLabel(item.source)),
      };
    }
    const result = aggregate(filteredRows, target, operation);
    if (result === undefined) {
      return {
        role: "assistant",
        text: `Encontré el campo ${target}, pero no contiene valores numéricos válidos para este cálculo.`,
        sources: source,
      };
    }
    const validCount = filteredRows.map(row => num(row[target])).filter(Number.isFinite).length;
    return {
      role: "assistant",
      text: `El ${operationLabel(operation).toLowerCase()} de ${target} es ${fmt(result)}, calculado con ${validCount} valores válidos${filterText(filters)}.`,
      table: [{ Indicador: `${operationLabel(operation)} de ${target}`, Valor: Number(result.toFixed(2)) }],
      sources: source,
    };
  }

  const meaningful = words(expanded).filter(word => !intentWords.has(word));
  const textualHits = ranked.flatMap(({ source: hitSource, score }) =>
    hitSource.rows
      .filter(row => meaningful.length && meaningful.some(term =>
        Object.values(row).some(value => clean(String(value ?? "")).includes(term))))
      .slice(0, 8)
      .map(row => ({ source: hitSource, score, row })))
    .sort((a, b) => b.score - a.score);

  if (wantsList || filters.length || textualHits.length) {
    const directRows = filters.length ? filteredRows.slice(0, 20) : [];
    if (directRows.length) {
      return {
        role: "assistant",
        text: `Encontré ${filteredRows.length} registros${filterText(filters)}. Muestro los primeros ${directRows.length}.`,
        table: directRows,
        sources: source,
      };
    }
    const hits = textualHits.slice(0, 12);
    if (hits.length) {
      return {
        role: "assistant",
        text: `Encontré ${textualHits.length} coincidencias relevantes en ${new Set(textualHits.map(hit => sourceLabel(hit.source))).size} hoja(s). Muestro las primeras ${hits.length}.`,
        table: hits.map(hit => ({
          Archivo: hit.source.workbook,
          Hoja: hit.source.sheet,
          ...hit.row,
        })),
        sources: [...new Set(hits.map(hit => sourceLabel(hit.source)))],
      };
    }
  }

  return {
    role: "assistant",
    text: `No encontré evidencia suficientemente directa para responder de forma confiable. ${coverage} Puedes escribir “¿Qué puedo preguntar?” para ver ejemplos adaptados a estos archivos.`,
    sources: relevant.slice(0, 3).map(sourceLabel),
  };
}

function chartKindFor(question: string, labelColumn: string, points: ChartPoint[]): ChartKind {
  const q = clean(question);
  if (["circular", "pastel", "torta", "composicion", "porcentaje", "participacion"].some(term => q.includes(term)) && points.length <= 10) {
    return "pie";
  }
  if (
    ["fecha", "dia", "semana", "mes", "hora"].some(term => clean(labelColumn).includes(term)) ||
    q.includes("evolucion") || q.includes("tendencia")
  ) {
    return "line";
  }
  return "bar";
}

export function attachChart(reply: Chat, question: string): Chat {
  if (reply.role !== "assistant") return reply;
  if (reply.chart?.length) {
    return {
      ...reply,
      chartKind: reply.chartKind ?? chartKindFor(question, reply.chartTitle ?? "Categoría", reply.chart),
      autoChart: reply.autoChart ?? operationRequested(question, "chart"),
    };
  }
  if (!reply.table?.length) return reply;
  const headers = Object.keys(reply.table[0] ?? {});
  const numericColumn = headers.find(header =>
    reply.table!.some(row => Number.isFinite(num(row[header]))));
  if (!numericColumn) return reply;
  const labelColumn = headers.find(header => header !== numericColumn) ?? headers[0];
  const chart = reply.table
    .map((row, index) => ({
      label: String(row[labelColumn] ?? `Dato ${index + 1}`),
      value: num(row[numericColumn]),
    }))
    .filter(point => Number.isFinite(point.value))
    .slice(0, 18);
  if (!chart.length) return reply;
  return {
    ...reply,
    chart,
    chartKind: chartKindFor(question, labelColumn, chart),
    chartTitle: `${numericColumn} por ${labelColumn}`,
    autoChart: operationRequested(question, "chart"),
  };
}

export const suggestedQuestions = [
  "¿Cuántas órdenes de trabajo hay?",
  "¿Cuántos técnicos tiene el taller?",
  "¿Cuántos lavadores hay?",
  "Gráfica de órdenes por estado",
  "Promedio de horas productivas por técnico",
  "Total de costo por tipo de vehículo",
  "¿Qué técnico tiene más horas productivas?",
  "¿Qué significa costo técnico?",
  "¿Qué puedo preguntar?",
];
