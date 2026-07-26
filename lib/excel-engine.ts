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

const normalizeForMatching = (s: string) =>
  clean(s)
    .replace(/\b(5|10|15|20|25|30|35|40|45|50|55)\s*mil\b/g, (_match, value: string) => `${value}000`)
    .replace(/(\d)[.,](?=\d{3}\b)/g, "$1")
    .replace(/\s+/g, " ")
    .trim();

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
    aliases: [
      "tipo de vehiculo", "tipo vehiculo", "clase de vehiculo", "categoria de vehiculo",
      "vehiculo", "vehiculos", "auto", "autos", "coche", "coches", "carro", "carros",
      "camioneta", "camionetas", "pickup", "pick up", "suv", "todoterreno",
      "automovil", "automoviles", "sedan", "sedanes", "monovolumen",
      "monovolumenes", "minivan", "minivanes",
    ],
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
    aliases: [
      "horas productivas", "hora productiva", "tiempo productivo", "productividad",
      "horas trabajadas", "horas trabajada", "horas trabajo", "horas que trabajo",
      "horas laboradas", "tiempo trabajado", "tiempo laborado", "carga horaria",
      "horas mensuales", "horas al mes",
      "trabajo mas horas", "trabajo menos horas", "mas horas trabajadas",
      "menos horas trabajadas", "tecnico que mas trabajo", "tecnico que menos trabajo",
      "quien trabajo mas", "quien trabajo menos",
    ],
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
    aliases: [
      "costo control calidad", "costo control de calidad",
      "costo de control de calidad", "costo calidad", "costo cc",
    ],
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
  count: [
    "cuanto", "cuantos", "cuanta", "cuantas", "cantidad", "numero de",
    "numero total de", "total de ordenes", "total de ot", "conteo", "cuenta",
    "existen", "hay", "tiene el taller",
  ],
  average: ["promedio", "media", "valor medio", "media aritmetica"],
  total: ["total", "suma", "sumatoria", "acumulado", "acumulada", "al mes", "mensual", "mensuales", "durante el mes", "en el mes"],
  max: ["maximo", "maxima", "mayor", "mas", "mas alto", "mas alta", "top"],
  min: ["minimo", "minima", "menor", "menos", "mas bajo", "mas baja"],
  list: ["muestra", "mostrar", "lista", "listar", "detalle", "detalla", "cuales", "quien", "quienes", "nombres", "registros", "que tipos", "que modelos", "que estados"],
  group: [" por ", " segun ", " cada ", "distribucion", "desglose", "composicion", "agrupa", "agrupado"],
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

const vehicleTypeRules = [
  {
    value: "Camioneta",
    singular: "camioneta",
    plural: "camionetas",
    aliases: ["camioneta", "camionetas", "pickup", "pickups", "pick up", "pick ups", "pick-up", "pick-ups"],
  },
  {
    value: "SUV",
    singular: "SUV",
    plural: "SUV",
    aliases: ["suv", "suvs", "todoterreno", "todoterrenos"],
  },
  {
    value: "Automóvil",
    singular: "automóvil",
    plural: "automóviles",
    aliases: ["automovil", "automoviles", "sedan", "sedanes"],
  },
  {
    value: "Monovolumen",
    singular: "monovolumen",
    plural: "monovolúmenes",
    aliases: ["monovolumen", "monovolumenes", "minivan", "minivanes"],
  },
];

function includesPhrase(question: string, phrase: string) {
  const normalizedQuestion = ` ${clean(question).replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ")} `;
  const normalizedPhrase = clean(phrase).replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
  return normalizedPhrase.length > 0 && normalizedQuestion.includes(` ${normalizedPhrase} `);
}

function detectVehicleType(question: string) {
  return vehicleTypeRules.find(rule =>
    rule.aliases.some(alias => includesPhrase(question, alias)));
}

export function expandQuestion(question: string) {
  const base = ` ${clean(question)} `;
  const additions = new Set<string>();
  for (const concept of concepts) {
    if (concept.aliases.some(alias => base.includes(` ${clean(alias)} `) || base.includes(clean(alias)))) {
      additions.add(concept.key);
    }
  }
  Object.entries(operationAliases).forEach(([operation, aliases]) => {
    if (aliases.some(alias => includesPhrase(base, alias))) additions.add(operation);
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
  if (
    query.includes("tipo vehiculo") &&
    source.workbook.includes("Entregable 3") &&
    clean(source.sheet) === "base_datos"
  ) {
    score += 30;
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
  const query = ` ${normalizeForMatching(question)} `;
  const queryStems = new Set(stems(query));
  const matches: { column: string; value: string; score: number }[] = [];
  for (const column of source.columns) {
    const values = [...new Set(
      source.rows
        .map(row => String(row[column] ?? "").trim())
        .filter(value => value && value.length < 90),
    )].slice(0, 500);
    for (const value of values) {
      const normalizedValue = normalizeForMatching(value);
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
  return operationAliases[operation].some(alias => includesPhrase(query, alias));
}

function wantsHelp(query: string) {
  const q = clean(query);
  return [
    "que puedo preguntar", "que preguntas", "preguntas disponibles", "ejemplos de preguntas",
    "ayuda", "como preguntar", "consultas disponibles", "que sabes",
  ].some(phrase => q.includes(phrase));
}

export type QuestionCategory = {
  title: string;
  description: string;
  questions: string[];
};

function uniqueQuestions(questions: string[]) {
  const seen = new Set<string>();
  return questions.filter(question => {
    const normalized = clean(question);
    if (!normalized || seen.has(normalized)) return false;
    seen.add(normalized);
    return true;
  });
}

function valuesFrom(
  source: SourceTable | undefined,
  columnNames: string[],
  limit = 40,
) {
  if (!source) return [];
  const column = source.columns.find(candidate =>
    columnNames.some(name => clean(candidate) === clean(name)));
  if (!column) return [];
  return [...new Set(
    source.rows
      .map(row => String(row[column] ?? "").trim())
      .filter(value => value && value.length < 80),
  )].slice(0, limit);
}

export function buildQuestionCatalog(sources: SourceTable[]): QuestionCategory[] {
  const dashboard = sources.find(source =>
    source.workbook.includes("Entregable 3") &&
    clean(source.sheet) === "base_datos");
  const personnel = sources.find(source => clean(source.sheet) === "personal");
  const dictionary = sources.find(source => clean(source.sheet) === "diccionario");
  const planning = sources.find(source =>
    source.workbook.includes("Entregable 2") &&
    clean(source.sheet) === "planning_diario");
  const capacity = sources.find(source => clean(source.sheet) === "capacidad_taller");

  const states = valuesFrom(dashboard, ["Estado"]);
  const vehicleTypes = valuesFrom(dashboard, ["Tipo vehículo", "Tipo de vehículo"]);
  const workTypes = valuesFrom(dashboard, ["Tipo trabajo", "Tipo mantenimiento"]);
  const situations = valuesFrom(dashboard, ["Situación"]);
  const compliance = valuesFrom(dashboard, ["Cumplimiento"]);
  const technicians = valuesFrom(dashboard, ["Técnico"]);
  const advisors = valuesFrom(dashboard, ["Asesor"]);
  const bays = valuesFrom(planning, ["Bahía"]);
  const weeks = valuesFrom(dashboard, ["Semana"]);
  const fields = valuesFrom(dictionary, ["Campo"]);

  const categories: QuestionCategory[] = [
    {
      title: "Órdenes de trabajo y estados",
      description: "Conteos, listados, estados y distribuciones de las OT.",
      questions: [
        "¿Cuántas órdenes de trabajo hay en el taller?",
        "¿Cuántas OT hay registradas?",
        "¿Cuál es el número total de órdenes?",
        "Cantidad de órdenes de trabajo",
        "Lista las órdenes de trabajo",
        "¿Qué estados de órdenes existen?",
        "Distribución de órdenes por estado",
        "Gráfica de órdenes por estado",
        "¿Qué estado tiene más órdenes?",
        "¿Qué estado tiene menos órdenes?",
        "Órdenes por técnico",
        "Órdenes por asesor",
        "Órdenes por fecha",
        "Órdenes por semana",
        "Órdenes por bahía",
        "Órdenes por tipo de vehículo",
        "Órdenes por tipo de trabajo",
        ...states.flatMap(state => [
          `¿Cuántas órdenes están ${state}?`,
          `¿Cuántas OT tienen estado ${state}?`,
          `Número de órdenes ${state}`,
          `Lista las órdenes ${state}`,
          `Órdenes ${state} por técnico`,
          `Gráfica de órdenes ${state} por tipo de vehículo`,
        ]),
      ],
    },
    {
      title: "Personal del taller",
      description: "Técnicos, asesores, lavado, control de calidad y personal completo.",
      questions: [
        "¿Cuántas personas trabajan en el taller?",
        "Lista todo el personal del taller",
        "Personal por cargo",
        "Gráfica del personal por cargo",
        "¿Cuántos técnicos tiene el taller?",
        "¿Cuántos mecánicos hay?",
        "Número de técnicos mecánicos",
        "¿Quiénes son los técnicos?",
        "Lista los mecánicos del taller",
        "¿Cuántos asesores de servicio hay?",
        "Número de recepcionistas",
        "¿Quiénes son los asesores?",
        "¿Cuántos lavadores hay?",
        "¿Cuánto personal de lavado existe?",
        "¿Quién es el lavador?",
        "¿Cuántas personas de control de calidad hay?",
        "Número de inspectores de calidad",
        "¿Quién está encargado del control de calidad?",
        ...technicians.flatMap(technician => [
          `¿Cuántas horas trabajó ${technician}?`,
          `Total de horas productivas de ${technician}`,
          `Total de horas extra de ${technician}`,
          `Costo técnico total de ${technician}`,
          `¿Cuántas órdenes tiene ${technician}?`,
          `Lista las órdenes de ${technician}`,
        ]),
        ...advisors.flatMap(advisor => [
          `¿Cuántas órdenes tiene el asesor ${advisor}?`,
          `Lista las órdenes de ${advisor}`,
          `Costo de asesor total de ${advisor}`,
        ]),
      ],
    },
    {
      title: "Vehículos",
      description: "Tipos de vehículo, modelos registrados y cruces con estados o servicios.",
      questions: [
        "¿Cuántos vehículos están registrados en el taller?",
        "¿Qué tipos de vehículos hay?",
        "Lista los tipos de vehículos registrados",
        "Distribución de vehículos por tipo",
        "Gráfica de vehículos por tipo",
        "¿Qué tipo de vehículo tiene más órdenes?",
        "¿Qué tipo de vehículo tiene menos órdenes?",
        "Vehículos por estado",
        "Vehículos por tipo de trabajo",
        ...vehicleTypes.flatMap(type => [
          `¿Cuántos vehículos de tipo ${type} están registrados en el taller?`,
          `¿Cuántos ${type} hay?`,
          `Número de vehículos ${type}`,
          `Lista los vehículos de tipo ${type}`,
          `${type} por estado`,
          `Gráfica de ${type} por tipo de trabajo`,
          `Costo total de vehículos ${type}`,
          `Promedio de horas productivas de vehículos ${type}`,
        ]),
      ],
    },
    {
      title: "Mantenimientos y servicios",
      description: "Tipos de trabajo y mantenimientos por kilometraje.",
      questions: [
        "¿Qué tipos de mantenimiento hay?",
        "Lista los tipos de trabajo",
        "Distribución de órdenes por tipo de trabajo",
        "Gráfica de mantenimientos por tipo de trabajo",
        "¿Qué mantenimiento tiene más órdenes?",
        "¿Qué mantenimiento tiene menos órdenes?",
        "Costo total por tipo de trabajo",
        "Promedio de horas productivas por tipo de trabajo",
        "Total de horas extra por tipo de trabajo",
        ...workTypes.flatMap(workType => {
          const mileage = normalizeForMatching(workType).match(/\b\d{4,6}\b/)?.[0];
          const shortMileage = mileage ? `${Number(mileage) / 1000} mil km` : workType;
          return [
            `¿Cuántas órdenes de ${workType} hay?`,
            `¿Cuántos mantenimientos de ${shortMileage} existen?`,
            `Número de OT para ${workType}`,
            `Lista las órdenes de ${workType}`,
            `${workType} por estado`,
            `${workType} por técnico`,
            `Costo total de ${workType}`,
            `Promedio de horas productivas de ${workType}`,
            `Gráfica de ${workType} por tipo de vehículo`,
          ];
        }),
      ],
    },
    {
      title: "Horas y productividad",
      description: "Horas productivas, requeridas, normales y extra, con totales y promedios.",
      questions: [
        "¿Cuántas horas trabajó al mes cada técnico en el taller?",
        "¿Cuántas horas trabajó cada técnico?",
        "Horas laboradas por técnico",
        "Total de horas productivas por técnico",
        "Promedio de horas productivas por técnico",
        "Gráfica de horas productivas por técnico",
        "¿Qué técnico trabajó más horas?",
        "¿Qué técnico trabajó menos horas?",
        "¿Quién acumula más horas productivas?",
        "¿Quién acumula menos horas productivas?",
        "Total de horas productivas del taller",
        "Promedio de horas productivas",
        "Total de horas productivas por fecha",
        "Total de horas productivas por semana",
        "Total de horas productivas por estado",
        "Total de horas productivas por tipo de vehículo",
        "Total de horas productivas por tipo de trabajo",
        "Horas extra acumuladas por técnico",
        "Total de horas extra por técnico",
        "Promedio de horas extra por técnico",
        "Total de horas extra por fecha",
        "Total de horas extra por semana",
        "Gráfica de horas extra por técnico",
        "Horas normales acumuladas por técnico",
        "Total de horas normales por técnico",
        "Promedio de horas normales por técnico",
        "Horas requeridas acumuladas por técnico",
        "Total de horas requeridas por técnico",
        "Promedio de horas requeridas por técnico",
        "Promedio del tempario por tipo de trabajo",
        "Máximo del tempario por tipo de vehículo",
        "Mínimo del tempario por tipo de trabajo",
        ...weeks.flatMap(week => [
          `Total de horas productivas de la semana ${week}`,
          `Horas productivas por técnico en la semana ${week}`,
          `Total de horas extra de la semana ${week}`,
        ]),
      ],
    },
    {
      title: "Costos, precios y ganancia",
      description: "Cálculos económicos por técnico, vehículo, trabajo, fecha o estado.",
      questions: [
        ...[
          "costo técnico", "costo asesor", "costo lavado",
          "costo control de calidad", "costo total", "precio estimado",
          "ganancia estimada",
        ].flatMap(metric => [
          `Total de ${metric}`,
          `Promedio de ${metric}`,
          `Máximo de ${metric}`,
          `Mínimo de ${metric}`,
          `Total de ${metric} por técnico`,
          `Total de ${metric} por fecha`,
          `Total de ${metric} por semana`,
          `Total de ${metric} por estado`,
          `Total de ${metric} por tipo de vehículo`,
          `Total de ${metric} por tipo de trabajo`,
          `Gráfica de ${metric} por técnico`,
          `Gráfica de ${metric} por tipo de vehículo`,
        ]),
        "¿Qué técnico acumula mayor costo técnico?",
        "¿Qué tipo de vehículo tiene mayor costo total?",
        "¿Qué tipo de trabajo genera mayor ganancia estimada?",
      ],
    },
    {
      title: "Capacidad y carga del taller",
      description: "Capacidad normal, programable, total, disponibilidad, carga y utilización.",
      questions: [
        ...[
          "capacidad normal", "capacidad programable", "capacidad total",
          "carga asignada", "disponible", "utilización",
        ].flatMap(metric => [
          `Total de ${metric}`,
          `Promedio de ${metric}`,
          `Máximo de ${metric}`,
          `Mínimo de ${metric}`,
          `${metric} por fecha`,
          `Gráfica de ${metric} por fecha`,
        ]),
        "¿Qué día tuvo mayor carga asignada?",
        "¿Qué día tuvo menor carga asignada?",
        "¿Qué día tuvo mayor utilización?",
        "¿Qué día tuvo más capacidad disponible?",
        "Distribución del diagnóstico de capacidad",
        "Gráfica del diagnóstico por fecha",
      ],
    },
    {
      title: "Planificación y recursos",
      description: "Distribuciones por técnico, asesor, bahía, fecha y semana.",
      questions: [
        "Órdenes por técnico",
        "Gráfica de órdenes por técnico",
        "Órdenes por asesor",
        "Gráfica de órdenes por asesor",
        "Órdenes por bahía",
        "Gráfica de órdenes por bahía",
        "Órdenes por fecha",
        "Gráfica de órdenes por fecha",
        "Órdenes por semana",
        "Gráfica de órdenes por semana",
        ...bays.flatMap(bay => [
          `¿Cuántas órdenes hay en la bahía ${bay}?`,
          `Lista las órdenes de la bahía ${bay}`,
          `Órdenes de la bahía ${bay} por estado`,
        ]),
        ...advisors.flatMap(advisor => [
          `Órdenes del asesor ${advisor} por estado`,
          `Órdenes del asesor ${advisor} por tipo de vehículo`,
        ]),
      ],
    },
    {
      title: "Cumplimiento y conflictos",
      description: "Retrasos, terminación, horas extra, conflictos y validez de la planificación.",
      questions: [
        "¿Cuántas órdenes tienen horas extra?",
        "¿Cuántas OT requieren sobretiempo?",
        "¿Cuántos conflictos de técnico existen?",
        "¿Cuántos solapamientos de técnicos hay?",
        "¿Cuántas órdenes están paradas?",
        "¿Cuántas entregas incumplidas hay?",
        "¿Cuántas órdenes válidas existen?",
        "Distribución de cumplimiento",
        "Gráfica de cumplimiento",
        "Cumplimiento por técnico",
        "Cumplimiento por tipo de vehículo",
        "Situación de las órdenes",
        "Gráfica de órdenes por situación",
        ...compliance.flatMap(value => [
          `¿Cuántas órdenes están ${value}?`,
          `Lista las órdenes con cumplimiento ${value}`,
          `Órdenes ${value} por técnico`,
        ]),
        ...situations.flatMap(value => [
          `¿Cuántas órdenes tienen situación ${value}?`,
          `Lista las órdenes ${value}`,
          `Órdenes ${value} por estado`,
        ]),
      ],
    },
    {
      title: "Definiciones de los Excel",
      description: "Significado y forma de cálculo de los campos documentados.",
      questions: fields.flatMap(field => [
        `¿Qué significa ${field}?`,
        `Define ${field}`,
        `¿Cómo se calcula ${field}?`,
      ]),
    },
  ];

  return categories
    .map(category => ({
      ...category,
      questions: uniqueQuestions(category.questions),
    }))
    .filter(category => category.questions.length);
}

function helpCatalog(sources: SourceTable[]): Row[] {
  return buildQuestionCatalog(sources).map(category => ({
    Tema: category.title,
    Preguntas: category.questions.length,
    Ejemplos: category.questions.slice(0, 3).join(" · "),
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

function periodForSource(source: SourceTable) {
  const dateColumn = source.columns.find(column =>
    ["fecha", "fecha programacion"].includes(clean(column)));
  if (!dateColumn) return undefined;
  const dates = source.rows
    .map(row => String(row[dateColumn] ?? "").trim())
    .map(value => {
      const match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (!match) return undefined;
      return {
        day: Number(match[1]),
        month: Number(match[2]),
        year: Number(match[3]),
        timestamp: Date.UTC(Number(match[3]), Number(match[2]) - 1, Number(match[1])),
      };
    })
    .filter((value): value is {
      day: number;
      month: number;
      year: number;
      timestamp: number;
    } => Boolean(value));
  if (!dates.length) return undefined;
  dates.sort((a, b) => a.timestamp - b.timestamp);
  const first = dates[0];
  const last = dates[dates.length - 1];
  const monthNames = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
  ];
  const sameMonth = first.month === last.month && first.year === last.year;
  return {
    first,
    last,
    text: sameMonth
      ? `${first.day} al ${last.day} de ${monthNames[first.month - 1]} de ${first.year}`
      : `${first.day}/${first.month}/${first.year} al ${last.day}/${last.month}/${last.year}`,
    isFullMonth: sameMonth &&
      first.day === 1 &&
      last.day === new Date(Date.UTC(last.year, last.month, 0)).getUTCDate(),
  };
}

function technicianHoursAnswer(
  question: string,
  sources: SourceTable[],
  wantsComputedTotal: boolean,
  wantsAverage: boolean,
  wantsMax: boolean,
  wantsMin: boolean,
) {
  const q = clean(question);
  const expanded = expandQuestion(question);
  const asksTechnicians = [
    "tecnico", "tecnicos", "mecanico", "mecanicos", "cada tecnico",
    "por tecnico", "personal tecnico",
  ].some(term => q.includes(term));
  const asksHours = [
    "horas productivas", "horas requeridas", "horas normales", "horas extra",
  ].some(key => expanded.includes(key));
  const naturalTotal = [
    "cada tecnico", "por tecnico", "los tecnicos", "tecnicos del taller",
    "horas trabajadas", "horas laboradas", "horas al mes", "horas mensuales",
  ].some(term => q.includes(term));
  if (
    !asksTechnicians ||
    !asksHours ||
    wantsAverage ||
    wantsMax ||
    wantsMin ||
    (!wantsComputedTotal && !naturalTotal)
  ) {
    return undefined;
  }

  const metricKey = expanded.includes("horas extra")
    ? "horas extra"
    : expanded.includes("horas normales")
      ? "horas normales"
      : expanded.includes("horas requeridas")
        ? "horas requeridas"
        : "horas productivas";
  const preferred = sources.find(source => {
    const hasTechnician = source.columns.some(column => clean(column) === "tecnico");
    const hasMetric = source.columns.some(column => clean(column) === metricKey);
    if (!hasTechnician || !hasMetric) return false;
    if (metricKey === "horas productivas" || metricKey === "horas extra") {
      return source.workbook.includes("Entregable 3") && clean(source.sheet) === "base_datos";
    }
    return source.workbook.includes("Entregable 2") && clean(source.sheet) === "planning_diario";
  }) ?? sources.find(source =>
    source.columns.some(column => clean(column) === "tecnico") &&
    source.columns.some(column => clean(column) === metricKey));
  if (!preferred) return undefined;

  const technicianColumn = preferred.columns.find(column => clean(column) === "tecnico")!;
  const metricColumn = preferred.columns.find(column => clean(column) === metricKey)!;
  const appliedFilters = filtersFor(question, preferred)
    .filter(filter => clean(filter.column) !== clean(metricColumn));
  const rows = applyFilters(preferred.rows, appliedFilters);
  const totals = new Map<string, number>();
  rows.forEach(row => {
    const technician = String(row[technicianColumn] ?? "").trim();
    const value = num(row[metricColumn]);
    if (technician && Number.isFinite(value)) {
      totals.set(technician, (totals.get(technician) ?? 0) + value);
    }
  });
  const personal = sources.find(source => clean(source.sheet) === "personal");
  const nameColumn = personal?.columns.find(column => clean(column) === "nombre");
  const preferredOrder = nameColumn
    ? personal!.rows.map(row => String(row[nameColumn] ?? "").trim())
    : [];
  const data = [...totals]
    .map(([label, value]) => ({ label, value: Number(value.toFixed(2)) }))
    .sort((a, b) => {
      const aIndex = preferredOrder.indexOf(a.label);
      const bIndex = preferredOrder.indexOf(b.label);
      return (aIndex < 0 ? 999 : aIndex) - (bIndex < 0 ? 999 : bIndex);
    });
  if (!data.length) return undefined;

  const period = periodForSource(preferred);
  const asksMonth = ["mes", "mensual", "mensuales", "julio"].some(term => q.includes(term));
  const periodText = period
    ? asksMonth && !period.isFullMonth
      ? `Los archivos no cubren el mes completo; contienen información del ${period.text}. Para ese período`
      : `Para el período del ${period.text}`
    : "Para el período disponible";
  return {
    role: "assistant" as const,
    text: `${periodText}, las ${metricColumn.toLowerCase()} acumuladas por técnico son: ${data.map(item => `${item.label}: ${fmt(item.value)} h`).join("; ")}.`,
    table: data.map(item => ({
      Técnico: item.label,
      [metricColumn]: item.value,
    })),
    chart: data,
    chartTitle: `${metricColumn} por técnico`,
    sources: [sourceLabel(preferred)],
  };
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

function vehicleTypeCountAnswer(
  question: string,
  sources: SourceTable[],
  wantsCount: boolean,
) {
  const vehicleType = detectVehicleType(question);
  if (!wantsCount || !vehicleType) return undefined;

  const candidates = sources
    .map(source => {
      const typeColumn = source.columns.find(column =>
        ["tipo vehiculo", "tipo de vehiculo"].includes(clean(column)));
      const idColumn = source.columns.find(column => clean(column) === "ot");
      if (!typeColumn || !idColumn) return undefined;
      const detectedFilters = filtersFor(question, source)
        .filter(filter => clean(filter.column) !== "ot");
      const extraFilters = detectedFilters
        .filter(filter => clean(filter.column) !== clean(typeColumn));
      const filteredRows = applyFilters(
        source.rows.filter(row =>
          clean(String(row[typeColumn] ?? "")) === clean(vehicleType.value)),
        extraFilters,
      );
      return {
        source,
        typeColumn,
        extraFilters,
        count: new Set(
          filteredRows
            .map(row => String(row[idColumn] ?? "").trim())
            .filter(Boolean),
        ).size,
      };
    })
    .filter((item): item is {
      source: SourceTable;
      typeColumn: string;
      extraFilters: AppliedFilter[];
      count: number;
    } => Boolean(item));

  if (!candidates.length) return undefined;

  const principal = candidates.find(candidate =>
    candidate.source.workbook.includes("Entregable 3") &&
    clean(candidate.source.sheet) === "base_datos") ?? candidates[0];
  const requiredExtraValues = principal.extraFilters.map(filter => clean(filter.value));
  const comparable = candidates.filter(candidate =>
    candidate === principal ||
    requiredExtraValues.every(value =>
      candidate.extraFilters.some(filter => clean(filter.value) === value)));
  const ordered = [
    principal,
    ...comparable.filter(candidate => candidate !== principal),
  ];
  const different = ordered.some(candidate => candidate.count !== principal.count);
  const extraText = principal.extraFilters.length
    ? ` con ${principal.extraFilters.map(filter => `${filter.column} = “${filter.value}”`).join(" y ")}`
    : "";
  const comparisonText = different
    ? ` En el control cruzado, ${ordered.slice(1).map(candidate =>
      `${candidate.source.workbook} registra ${candidate.count}`).join("; ")}. Los archivos no coinciden en este indicador.`
    : ordered.length > 1
      ? ` Las ${ordered.length} fuentes detalladas coinciden.`
      : "";
  const vehicleWord = principal.count === 1 ? "vehículo" : "vehículos";
  const registeredWord = principal.count === 1 ? "registrado" : "registrados";

  return {
    role: "assistant" as const,
    text: `Hay ${principal.count} ${vehicleWord} de tipo ${vehicleType.singular}${extraText} ${registeredWord} en el taller, según el Dashboard de indicadores, que se usa como fuente principal para esta consulta.${comparisonText}`,
    table: ordered.map(candidate => ({
      Archivo: candidate.source.workbook,
      Hoja: candidate.source.sheet,
      "Tipo de vehículo": vehicleType.value,
      Cantidad: candidate.count,
      Uso: candidate === principal ? "Valor principal" : "Control cruzado",
    })),
    chart: ordered.map(candidate => ({
      label: candidate === principal
        ? `Dashboard: ${vehicleType.plural}`
        : `${candidate.source.workbook.replace(/ · .*/, "")}: ${vehicleType.plural}`,
      value: candidate.count,
    })),
    chartTitle: `${vehicleType.plural} registradas por fuente`,
    sources: ordered.map(candidate => sourceLabel(candidate.source)),
  };
}

function vehicleTypeListAnswer(
  question: string,
  sources: SourceTable[],
  wantsList: boolean,
) {
  const q = clean(question);
  const asksTypes = q.includes("tipo") || q.includes("categoria") || q.includes("clase");
  const asksVehicles = ["vehiculo", "vehiculos", "carro", "carros", "auto", "autos"].some(term =>
    q.includes(term));
  if (!wantsList || !asksTypes || !asksVehicles || detectVehicleType(question)) return undefined;

  const principal = sources.find(source =>
    source.workbook.includes("Entregable 3") &&
    clean(source.sheet) === "base_datos" &&
    source.columns.some(column => ["tipo vehiculo", "tipo de vehiculo"].includes(clean(column))));
  if (!principal) return undefined;
  const typeColumn = principal.columns.find(column =>
    ["tipo vehiculo", "tipo de vehiculo"].includes(clean(column)))!;
  const idColumn = principal.columns.find(column => clean(column) === "ot");
  const buckets = new Map<string, Set<string>>();
  principal.rows.forEach((row, index) => {
    const label = String(row[typeColumn] ?? "").trim();
    if (!label) return;
    if (!buckets.has(label)) buckets.set(label, new Set());
    buckets.get(label)!.add(idColumn ? String(row[idColumn] ?? index) : String(index));
  });
  const data = [...buckets]
    .map(([label, values]) => ({ label, value: values.size }))
    .sort((a, b) => b.value - a.value);
  return {
    role: "assistant" as const,
    text: `El Dashboard registra ${data.length} tipos de vehículos: ${data.map(item => `${item.label} (${item.value})`).join(", ")}.`,
    table: data.map(item => ({
      "Tipo de vehículo": item.label,
      Cantidad: item.value,
    })),
    chart: data,
    chartTitle: "Vehículos registrados por tipo",
    sources: [sourceLabel(principal)],
  };
}

function workTypeCountAnswer(
  question: string,
  sources: SourceTable[],
  wantsCount: boolean,
) {
  if (!wantsCount) return undefined;
  const principalSource = sources.find(source =>
    source.workbook.includes("Entregable 3") &&
    clean(source.sheet) === "base_datos");
  if (!principalSource) return undefined;
  const principalColumn = principalSource.columns.find(column =>
    ["tipo trabajo", "tipo mantenimiento", "tipo de mantenimiento"].includes(clean(column)));
  const principalId = principalSource.columns.find(column => clean(column) === "ot");
  if (!principalColumn || !principalId) return undefined;
  const principalWorkFilter = filtersFor(question, principalSource)
    .find(filter => clean(filter.column) === clean(principalColumn));
  if (!principalWorkFilter) return undefined;

  const candidates = sources
    .map(source => {
      const workColumn = source.columns.find(column =>
        ["tipo trabajo", "tipo mantenimiento", "tipo de mantenimiento"].includes(clean(column)));
      const idColumn = source.columns.find(column => clean(column) === "ot");
      if (!workColumn || !idColumn) return undefined;
      const principalWork = normalizeForMatching(principalWorkFilter.value);
      const principalMileage = principalWork.match(/\b\d{4,6}\b/)?.[0];
      const matchingRows = source.rows.filter(row => {
        const candidateWork = normalizeForMatching(String(row[workColumn] ?? ""));
        return candidateWork === principalWork || Boolean(
          principalMileage &&
          candidateWork.includes("mantenimiento") &&
          candidateWork.includes(principalMileage),
        );
      });
      if (!matchingRows.length) return undefined;
      const extraFilters = filtersFor(question, source)
        .filter(filter =>
          clean(filter.column) !== clean(workColumn) &&
          clean(filter.column) !== "ot");
      const filteredRows = applyFilters(matchingRows, extraFilters);
      return {
        source,
        extraFilters,
        count: new Set(
          filteredRows.map(row => String(row[idColumn] ?? "").trim()).filter(Boolean),
        ).size,
      };
    })
    .filter((item): item is {
      source: SourceTable;
      extraFilters: AppliedFilter[];
      count: number;
    } => Boolean(item));
  if (!candidates.length) return undefined;

  const principal = candidates.find(candidate => candidate.source === principalSource) ?? candidates[0];
  const ordered = [principal, ...candidates.filter(candidate => candidate !== principal)];
  const different = ordered.some(candidate => candidate.count !== principal.count);
  const extraText = principal.extraFilters.length
    ? ` con ${principal.extraFilters.map(filter => `${filter.column} = “${filter.value}”`).join(" y ")}`
    : "";
  const comparisonText = different
    ? ` En el control cruzado, ${ordered.slice(1).map(candidate =>
      `${candidate.source.workbook} registra ${candidate.count}`).join("; ")}.`
    : ordered.length > 1
      ? ` Las ${ordered.length} fuentes detalladas coinciden.`
      : "";
  return {
    role: "assistant" as const,
    text: `Hay ${principal.count} órdenes de “${principalWorkFilter.value}”${extraText}, según el Dashboard de indicadores.${comparisonText}`,
    table: ordered.map(candidate => ({
      Archivo: candidate.source.workbook,
      Hoja: candidate.source.sheet,
      "Tipo de trabajo": principalWorkFilter.value,
      Cantidad: candidate.count,
      Uso: candidate === principal ? "Valor principal" : "Control cruzado",
    })),
    chart: ordered.map(candidate => ({
      label: candidate === principal ? "Dashboard" : candidate.source.workbook.replace(/ · .*/, ""),
      value: candidate.count,
    })),
    chartTitle: `Órdenes de ${principalWorkFilter.value}`,
    sources: ordered.map(candidate => sourceLabel(candidate.source)),
  };
}

function workTypeListAnswer(
  question: string,
  sources: SourceTable[],
  wantsList: boolean,
) {
  const q = clean(question);
  const asksTypes = q.includes("tipo") || q.includes("clase") || q.includes("categoria");
  const asksWork = ["trabajo", "mantenimiento", "servicio"].some(term => q.includes(term));
  if (!wantsList || !asksTypes || !asksWork) return undefined;
  const principal = sources.find(source =>
    source.workbook.includes("Entregable 3") &&
    clean(source.sheet) === "base_datos");
  if (!principal) return undefined;
  const workColumn = principal.columns.find(column => clean(column) === "tipo trabajo");
  const idColumn = principal.columns.find(column => clean(column) === "ot");
  if (!workColumn) return undefined;
  const buckets = new Map<string, Set<string>>();
  principal.rows.forEach((row, index) => {
    const label = String(row[workColumn] ?? "").trim();
    if (!label) return;
    if (!buckets.has(label)) buckets.set(label, new Set());
    buckets.get(label)!.add(idColumn ? String(row[idColumn] ?? index) : String(index));
  });
  const data = [...buckets]
    .map(([label, values]) => ({ label, value: values.size }))
    .sort((a, b) => b.value - a.value);
  return {
    role: "assistant" as const,
    text: `El Dashboard registra ${data.length} tipos de trabajo: ${data.map(item => `${item.label} (${item.value})`).join(", ")}.`,
    table: data.map(item => ({ "Tipo de trabajo": item.label, Cantidad: item.value })),
    chart: data,
    chartTitle: "Órdenes por tipo de trabajo",
    sources: [sourceLabel(principal)],
  };
}

function conflictIndicatorAnswer(
  question: string,
  sources: SourceTable[],
  wantsCount: boolean,
) {
  if (!wantsCount) return undefined;
  const q = clean(question);
  const rules = [
    {
      aliases: ["ordenes con horas extra", "ot con horas extra", "ordenes que requieren horas extra", "ordenes con sobretiempo"],
      indicator: "Órdenes con horas extra",
    },
    {
      aliases: ["conflictos de tecnico", "conflictos de tecnicos", "solapamientos de tecnico", "cruces de tecnico"],
      indicator: "Conflictos de técnico",
    },
    {
      aliases: ["entregas incumplidas", "ordenes incumplidas", "entregas fuera de plazo"],
      indicator: "Entregas incumplidas",
    },
    {
      aliases: ["ordenes validas", "ot validas", "ordenes sin conflicto"],
      indicator: "Órdenes válidas",
    },
  ];
  const rule = rules.find(candidate =>
    candidate.aliases.some(alias => q.includes(clean(alias))));
  if (!rule) return undefined;
  const source = sources.find(candidate =>
    clean(candidate.sheet) === "conflictos" &&
    candidate.columns.some(column => clean(column) === "indicador") &&
    candidate.columns.some(column => clean(column) === "cantidad"));
  if (!source) return undefined;
  const indicatorColumn = source.columns.find(column => clean(column) === "indicador")!;
  const countColumn = source.columns.find(column => clean(column) === "cantidad")!;
  const row = source.rows.find(candidate =>
    clean(String(candidate[indicatorColumn] ?? "")) === clean(rule.indicator));
  if (!row) return undefined;
  const count = num(row[countColumn]);
  if (!Number.isFinite(count)) return undefined;
  return {
    role: "assistant" as const,
    text: `Hay ${fmt(count)} ${rule.indicator.toLowerCase()}, según el resumen de conflictos del planning.`,
    table: [{ Indicador: rule.indicator, Cantidad: count }],
    sources: [sourceLabel(source)],
  };
}

function operationalCount(question: string, sources: SourceTable[], wantsCount: boolean) {
  const q = clean(question);
  const asksOrders = /\bot\b/.test(q) || q.includes("orden de trabajo") || q.includes("ordenes");
  const asksVehicles = q.includes("vehiculo") || q.includes("autos") || q.includes("coches") || q.includes("carros");
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
    const principal = comparable.find(item =>
      item.source.workbook.includes("Entregable 3") &&
      clean(item.source.sheet) === "base_datos");
    if (principal) {
      const ordered = [principal, ...comparable.filter(item => item !== principal)];
      return {
        role: "assistant" as const,
        text: `Según el Dashboard de indicadores, hay ${principal.count} ${entity}${filterText(principal.filters)}. Las fuentes operativas presentan valores diferentes; se muestra el control cruzado.`,
        table: ordered.map(item => ({
          Archivo: item.source.workbook,
          Hoja: item.source.sheet,
          Cantidad: item.count,
          Uso: item === principal ? "Valor principal" : "Control cruzado",
        })),
        sources: ordered.map(item => sourceLabel(item.source)),
      };
    }
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
  const numericConceptKeys = [
    "horas productivas", "horas requeridas", "horas normales", "horas extra",
    "tempario", "costo tecnico", "costo asesor", "costo lavado",
    "costo control calidad", "costo total", "precio estimado", "ganancia estimada",
    "capacidad", "utilizacion",
  ];
  const hasNumericConcept = numericConceptKeys.some(key => expanded.includes(key));
  const wantsNumericQuantity = wantsCount && hasNumericConcept;
  const wantsEntityCount = wantsCount && !wantsNumericQuantity;
  const wantsComputedTotal = wantsTotal || wantsNumericQuantity;

  if (wantsHelp(q)) {
    const catalog = buildQuestionCatalog(sources);
    const questionCount = catalog.reduce(
      (total, category) => total + category.questions.length,
      0,
    );
    return {
      role: "assistant",
      text: `Preparé ${questionCount} preguntas y variantes basadas en los valores reales de los cuatro Excel. Puedes preguntar con lenguaje natural: reconozco singular, plural, sinónimos, abreviaturas como OT y órdenes como contar, listar, filtrar, sumar, promediar, comparar o graficar.`,
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

  const technicianHours = technicianHoursAnswer(
    question,
    sources,
    wantsComputedTotal,
    wantsAverage,
    wantsMax,
    wantsMin,
  );
  if (technicianHours) return technicianHours;

  const conflictIndicator = conflictIndicatorAnswer(question, sources, wantsCount);
  if (conflictIndicator) return conflictIndicator;

  const personnel = personnelAnswer(question, sources, wantsEntityCount, wantsList);
  if (personnel) return personnel;

  const vehicleTypeCount = vehicleTypeCountAnswer(question, sources, wantsEntityCount);
  if (vehicleTypeCount && !wantsGroup) return vehicleTypeCount;

  const vehicleTypeList = vehicleTypeListAnswer(question, sources, wantsList);
  if (vehicleTypeList) return vehicleTypeList;

  const workTypeCount = workTypeCountAnswer(question, sources, wantsEntityCount);
  if (workTypeCount && !wantsGroup) return workTypeCount;

  const workTypeList = workTypeListAnswer(question, sources, wantsList);
  if (workTypeList) return workTypeList;

  const directCount = operationalCount(question, sources, wantsEntityCount);
  if (directCount && !wantsGroup) return directCount;

  const ranked = sources
    .map(source => ({ source, score: sourceScore(expanded, source) }))
    .sort((a, b) => b.score - a.score);
  const relevant = ranked.filter(item => item.score > 3).map(item => item.source);
  const pool = relevant.length ? relevant : sources;
  const needsNumericTarget =
    wantsAverage || wantsComputedTotal ||
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
  if (wantsEntityCount || wantsList || wantsChart) {
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

  const defaultGroupedCount = group && (
    wantsEntityCount ||
    (wantsGroup && !target) ||
    (wantsChart && (!target || asksOrders)) ||
    ((wantsMax || wantsMin) && asksOrders)
  );
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
        ["base_datos", "ordenes_trabajo", "planning_diario", "planning_ia"].includes(clean(candidate.source.sheet)))
      .sort((a, b) => {
        if (groupConcept !== "tipo vehiculo") return 0;
        const aPrincipal = a.source.workbook.includes("Entregable 3") && clean(a.source.sheet) === "base_datos" ? 1 : 0;
        const bPrincipal = b.source.workbook.includes("Entregable 3") && clean(b.source.sheet) === "base_datos" ? 1 : 0;
        return bPrincipal - aPrincipal;
      });
    const expectedFilterValues = filters
      .filter(filter => clean(filter.column) !== clean(group))
      .map(filter => clean(filter.value));
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
        candidateFilters,
        data: [...buckets]
          .map(([label, ids]) => ({ label, value: ids.size }))
          .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label)),
      };
    }).filter(distribution =>
      expectedFilterValues.every(value =>
        distribution.candidateFilters.some(filter => clean(filter.value) === value)));
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
      const principalDistribution = distributions.find(distribution =>
        distribution.source.workbook.includes("Entregable 3") &&
        clean(distribution.source.sheet) === "base_datos") ?? distributions[0];
      if (groupConcept === "tipo vehiculo" && principalDistribution) {
        return {
          role: "assistant",
          text: `Según el Dashboard de indicadores, la distribución principal es: ${principalDistribution.data.map(item => `${item.label} (${item.value})`).join(", ")}. El otro archivo detallado presenta valores diferentes; el desglose completo se muestra como control cruzado.`,
          table,
          chart: principalDistribution.data,
          chartTitle: "Distribución principal de vehículos por tipo",
          autoChart: wantsChart,
          sources: distributions.map(distribution => sourceLabel(distribution.source)),
        };
      }
      const requestedVehicleType = detectVehicleType(question);
      if (requestedVehicleType && principalDistribution) {
        const orderedDistributions = [
          principalDistribution,
          ...distributions.filter(distribution => distribution !== principalDistribution),
        ];
        const principalFirstTable = orderedDistributions.flatMap(distribution =>
          distribution.data.map(item => ({
            Archivo: distribution.source.workbook,
            Hoja: distribution.source.sheet,
            [group]: item.label,
            Cantidad: item.value,
          })));
        return {
          role: "assistant",
          text: `Según el Dashboard de indicadores, los vehículos de tipo ${requestedVehicleType.singular} se distribuyen por ${group} así: ${principalDistribution.data.map(item => `${item.label} (${item.value})`).join(", ")}. La otra fuente detallada presenta valores diferentes y se incluye como control cruzado.`,
          table: principalFirstTable,
          chart: principalDistribution.data,
          chartTitle: `${requestedVehicleType.plural} por ${group}`,
          autoChart: wantsChart,
          sources: orderedDistributions.map(distribution => sourceLabel(distribution.source)),
        };
      }
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

  if (wantsEntityCount) {
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
      : wantsComputedTotal ? "total"
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
  "¿Cuántas camionetas están registradas?",
  "¿Cuántas horas trabajó al mes cada técnico?",
  "Gráfica de órdenes por estado",
  "¿Qué técnico trabajó más horas?",
  "Total de costo por tipo de vehículo",
  "¿Cuántos lavadores hay?",
  "¿Qué significa costo técnico?",
  "¿Qué puedo preguntar?",
];
