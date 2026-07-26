"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
  answerAcross,
  attachChart,
  buildQuestionCatalog,
  ChartKind,
  ChartPoint,
  Chat,
  fmt,
  parseWorkbook,
  SourceTable,
  suggestedQuestions,
} from "../lib/excel-engine";

const databases = [
  { name: "Entregable 1 · Base de datos", file: "Planning_Carga_Trabajo_Entregable_1.xlsx" },
  { name: "Planning generado por IA", file: "Planning_IA.xlsx" },
  { name: "Entregable 3 · Dashboard", file: "Entregable_3_Dashboard_Indicadores.xlsx" },
  { name: "Entregable 2 · Planning", file: "Entregable_2_Planning_Carga_Trabajo.xlsx" },
];

function BarChart({ data }: { data: ChartPoint[] }) {
  const max = Math.max(...data.map(point => Math.abs(point.value)), 1);
  return (
    <div className="barChart">
      {data.map((point, index) => (
        <div className="barRow" key={`${point.label}-${index}`}>
          <span title={point.label}>{point.label}</span>
          <div>
            <i style={{ width: `${Math.max(4, (Math.abs(point.value) / max) * 100)}%` }} />
          </div>
          <b>{fmt(point.value)}</b>
        </div>
      ))}
    </div>
  );
}

function LineChart({ data }: { data: ChartPoint[] }) {
  const width = 680;
  const height = 230;
  const left = 42;
  const right = 18;
  const top = 18;
  const bottom = 45;
  const values = data.map(point => point.value);
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 1);
  const range = max - min || 1;
  const x = (index: number) =>
    left + (index * (width - left - right)) / Math.max(1, data.length - 1);
  const y = (value: number) =>
    top + ((max - value) * (height - top - bottom)) / range;
  const points = data.map((point, index) => `${x(index)},${y(point.value)}`).join(" ");
  return (
    <div className="lineChart">
      <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Gráfica de línea">
        {[0, 1, 2, 3, 4].map(index => {
          const gy = top + (index * (height - top - bottom)) / 4;
          const label = max - (index * range) / 4;
          return (
            <g key={index}>
              <line x1={left} x2={width - right} y1={gy} y2={gy} className="gridLine" />
              <text x={left - 7} y={gy + 4} textAnchor="end">{fmt(label)}</text>
            </g>
          );
        })}
        <polyline points={points} className="trendLine" />
        {data.map((point, index) => (
          <g key={`${point.label}-${index}`}>
            <circle cx={x(index)} cy={y(point.value)} r="4.5" />
            {(data.length <= 8 || index === 0 || index === data.length - 1) && (
              <text className="axisLabel" x={x(index)} y={height - 15} textAnchor="middle">
                {point.label.length > 12 ? `${point.label.slice(0, 11)}…` : point.label}
              </text>
            )}
          </g>
        ))}
      </svg>
    </div>
  );
}

const pieColors = [
  "#1f7a52",
  "#3d9d6d",
  "#72bd92",
  "#a5d6b9",
  "#d0b867",
  "#b98254",
  "#8f729f",
  "#567f9c",
  "#7d9688",
  "#294f40",
];

function PieChart({ data }: { data: ChartPoint[] }) {
  const visible = data.slice(0, 10);
  const total = visible.reduce((sum, point) => sum + Math.abs(point.value), 0) || 1;
  let cursor = 0;
  const gradient = visible.map((point, index) => {
    const start = cursor;
    cursor += (Math.abs(point.value) / total) * 100;
    return `${pieColors[index % pieColors.length]} ${start}% ${cursor}%`;
  }).join(", ");
  return (
    <div className="pieChart">
      <div className="donut" style={{ background: `conic-gradient(${gradient})` }}>
        <div><b>{fmt(total)}</b><small>Total</small></div>
      </div>
      <div className="legend">
        {visible.map((point, index) => (
          <div key={`${point.label}-${index}`}>
            <i style={{ background: pieColors[index % pieColors.length] }} />
            <span title={point.label}>{point.label}</span>
            <b>{((Math.abs(point.value) / total) * 100).toFixed(1)}%</b>
          </div>
        ))}
      </div>
    </div>
  );
}

function ChartView({ data, kind }: { data: ChartPoint[]; kind: ChartKind }) {
  if (kind === "line") return <LineChart data={data} />;
  if (kind === "pie") return <PieChart data={data} />;
  return <BarChart data={data} />;
}

export default function Home() {
  const [sources, setSources] = useState<SourceTable[]>([]);
  const [question, setQuestion] = useState("");
  const [chat, setChat] = useState<Chat[]>([{
    role: "assistant",
    text: "Estoy cargando e indexando los cuatro archivos Excel. Después podrás preguntar sin seleccionar un archivo.",
  }]);
  const [loading, setLoading] = useState(true);
  const [drag, setDrag] = useState(false);
  const [showCatalog, setShowCatalog] = useState(false);
  const [catalogSearch, setCatalogSearch] = useState("");
  const [visibleCharts, setVisibleCharts] = useState<Record<number, boolean>>({});
  const [chartKinds, setChartKinds] = useState<Record<number, ChartKind>>({});
  const inputRef = useRef<HTMLInputElement>(null);
  const totalRows = useMemo(
    () => sources.reduce((total, source) => total + source.rows.length, 0),
    [sources],
  );
  const questionCatalog = useMemo(() => buildQuestionCatalog(sources), [sources]);
  const questionCount = useMemo(
    () => questionCatalog.reduce((total, category) => total + category.questions.length, 0),
    [questionCatalog],
  );
  const filteredCatalog = useMemo(() => {
    const search = catalogSearch
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
    if (!search) return questionCatalog;
    return questionCatalog
      .map(category => ({
        ...category,
        questions: category.questions.filter(item =>
          `${category.title} ${item}`
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .includes(search)),
      }))
      .filter(category => category.questions.length);
  }, [catalogSearch, questionCatalog]);

  async function loadAll() {
    setLoading(true);
    try {
      const loaded = (await Promise.all(databases.map(async database => {
        const response = await fetch(`/data/${database.file}`);
        if (!response.ok) throw new Error(`No se pudo abrir ${database.file}`);
        return parseWorkbook(await response.arrayBuffer(), database.name, database.file);
      }))).flat();
      setSources(loaded);
      setChat([{
        role: "assistant",
        text: `Listo. Indexé los cuatro libros, ${loaded.length} hojas y ${loaded.reduce((total, source) => total + source.rows.length, 0)} registros. Puedes usar sinónimos, filtros, comparaciones y solicitudes de gráficas.`,
      }]);
    } catch (error) {
      setChat([{
        role: "assistant",
        text: `No pude cargar todos los Excel: ${error instanceof Error ? error.message : "error desconocido"}. Verifica que la carpeta public/data se haya subido completa.`,
      }]);
    } finally {
      setLoading(false);
    }
  }

  async function addFile(file: File) {
    try {
      const extra = parseWorkbook(await file.arrayBuffer(), file.name, file.name);
      setSources(current => [...current, ...extra]);
      setChat(current => [...current, {
        role: "assistant",
        text: `Agregué “${file.name}” al análisis conjunto con ${extra.length} hoja(s).`,
      }]);
    } catch {
      setChat(current => [...current, {
        role: "assistant",
        text: `No pude interpretar “${file.name}”. Verifica que sea un archivo Excel o CSV válido.`,
      }]);
    }
  }

  function ask(customQuestion?: string) {
    const value = (customQuestion ?? question).trim();
    if (!value || loading) return;
    const user: Chat = { role: "user", text: value };
    const reply = attachChart(answerAcross(value, sources), value);
    setChat(current => [...current, user, reply]);
    setQuestion("");
    setShowCatalog(false);
  }

  useEffect(() => {
    loadAll();
  }, []);

  return (
    <main>
      <header className="topbar">
        <div className="brand">
          <span className="logo">X</span>
          <div><strong>ExcelQ</strong><small>Asistente de datos</small></div>
        </div>
        <span className="privacy">Consulta conjunta · 4 libros Excel · Sin consumo de API</span>
      </header>

      <section className="hero">
        <div>
          <p className="eyebrow">BUSCA · RELACIONA · RESPONDE · GRAFICA</p>
          <h1>
            Universidad de las Fuerzas Armadas ESPE Latacunga
            <br />
            Mantenimiento Automotriz
          </h1>
          <p>
            El asistente interpreta variantes de preguntas, cruza los cuatro libros,
            muestra sus fuentes y genera gráficas sin usar una API de pago.
          </p>
        </div>
        <div className="metric">
          <span>{loading ? "…" : totalRows}</span>
          <small>registros indexados</small>
        </div>
      </section>

      <div className="workspace">
        <aside className="sidebar">
          <div className="indexTitle">
            <span className={loading ? "pulse" : ""} />
            <div>
              <strong>{loading ? "Indexando archivos" : "4 archivos conectados"}</strong>
              <small>{sources.length} hojas disponibles</small>
            </div>
          </div>
          <div className="fileList">
            {databases.map(database => (
              <div className="sourceFile" key={database.file}>
                <b>✓</b><span>{database.name}</span>
              </div>
            ))}
          </div>
          <div
            className={`drop compact ${drag ? "drag" : ""}`}
            onDragOver={event => { event.preventDefault(); setDrag(true); }}
            onDragLeave={() => setDrag(false)}
            onDrop={event => {
              event.preventDefault();
              setDrag(false);
              const file = event.dataTransfer.files[0];
              if (file) addFile(file);
            }}
          >
            <div className="fileIcon">+</div>
            <strong>Agregar otro Excel</strong>
            <p>Se consultará junto con los cuatro archivos</p>
            <button onClick={() => inputRef.current?.click()}>Seleccionar</button>
            <input
              ref={inputRef}
              hidden
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={event => event.target.files?.[0] && addFile(event.target.files[0])}
            />
          </div>
          <div className="guardrail">
            <b>Respuesta responsable</b>
            <p>Si el dato no existe o las fuentes no coinciden, el asistente lo indica y no inventa una cifra.</p>
          </div>
        </aside>

        <section className="chatPanel">
          <div className="chatHead">
            <div>
              <strong>Asistente de análisis integral</strong>
              <span className="online">{loading ? "Cargando" : "Listo"}</span>
            </div>
            <small>Busca e interpreta todos los archivos automáticamente</small>
          </div>
          <div className="questionTools">
            <div className="suggestionStrip" aria-label="Preguntas sugeridas">
              {suggestedQuestions.slice(0, 6).map(example => (
                <button key={example} onClick={() => ask(example)} disabled={loading}>
                  {example}
                </button>
              ))}
              <button
                className="catalogOpen"
                onClick={() => setShowCatalog(current => !current)}
                disabled={loading}
                aria-expanded={showCatalog}
              >
                {showCatalog ? "Cerrar preguntas" : `Ver todas (${questionCount})`}
              </button>
            </div>
            {showCatalog && (
              <section className="questionCatalog" aria-label="Catálogo de preguntas">
                <div className="catalogHead">
                  <div>
                    <strong>Preguntas preparadas desde los Excel</strong>
                    <small>{questionCount} consultas y variantes en {questionCatalog.length} temas</small>
                  </div>
                  <button onClick={() => setShowCatalog(false)} aria-label="Cerrar catálogo">×</button>
                </div>
                <input
                  className="catalogSearch"
                  value={catalogSearch}
                  onChange={event => setCatalogSearch(event.target.value)}
                  placeholder="Buscar: horas, camionetas, costos, técnico…"
                  autoFocus
                />
                <div className="catalogContent">
                  {filteredCatalog.map(category => (
                    <details key={category.title} open={Boolean(catalogSearch)}>
                      <summary>
                        <span>{category.title}</span>
                        <b>{category.questions.length}</b>
                      </summary>
                      <p>{category.description}</p>
                      <div className="catalogQuestions">
                        {category.questions.map(example => (
                          <button key={example} onClick={() => ask(example)}>
                            {example}
                          </button>
                        ))}
                      </div>
                    </details>
                  ))}
                  {!filteredCatalog.length && (
                    <p className="catalogEmpty">
                      No hay una pregunta predefinida con ese texto. Puedes escribirla directamente en el chat.
                    </p>
                  )}
                </div>
              </section>
            )}
          </div>
          <div className="messages">
            {chat.map((message, index) => {
              const chartVisible = visibleCharts[index] ?? Boolean(message.autoChart);
              const chartKind = chartKinds[index] ?? message.chartKind ?? "bar";
              return (
                <article key={index} className={`message ${message.role}`}>
                  <div className="avatar">{message.role === "assistant" ? "EQ" : "TÚ"}</div>
                  <div className="bubble">
                    <p>{message.text}</p>
                    {message.chart && (
                      <div className="chartBlock">
                        <div className="chartActions">
                          <div>
                            <button
                              className="chartToggle"
                              onClick={() => setVisibleCharts(current => ({
                                ...current,
                                [index]: !chartVisible,
                              }))}
                            >
                              {chartVisible ? "Ocultar gráfica" : "Generar gráfica"}
                            </button>
                            {chartVisible && (
                              <span className="chartTitle">{message.chartTitle}</span>
                            )}
                          </div>
                          {chartVisible && (
                            <div className="chartTypes" aria-label="Tipo de gráfica">
                              {([
                                ["bar", "Barras"],
                                ["line", "Línea"],
                                ["pie", "Circular"],
                              ] as [ChartKind, string][]).map(([kind, label]) => (
                                <button
                                  key={kind}
                                  className={chartKind === kind ? "active" : ""}
                                  onClick={() => setChartKinds(current => ({
                                    ...current,
                                    [index]: kind,
                                  }))}
                                >
                                  {label}
                                </button>
                              ))}
                            </div>
                          )}
                        </div>
                        {chartVisible && <ChartView data={message.chart} kind={chartKind} />}
                      </div>
                    )}
                    {message.table && (
                      <div className="tableWrap">
                        <table>
                          <thead>
                            <tr>
                              {Object.keys(message.table[0] ?? {}).map(header => (
                                <th key={header}>{header}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {message.table.slice(0, 20).map((row, rowIndex) => (
                              <tr key={rowIndex}>
                                {Object.keys(message.table![0] ?? {}).map(header => (
                                  <td key={header}>{String(row[header] ?? "")}</td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                    {message.sources && message.sources.length > 0 && (
                      <div className="sources">
                        <b>Fuentes consultadas:</b> {message.sources.join(" · ")}
                      </div>
                    )}
                  </div>
                </article>
              );
            })}
          </div>
          <div className="composer">
            <input
              value={question}
              onChange={event => setQuestion(event.target.value)}
              onKeyDown={event => event.key === "Enter" && ask()}
              placeholder={loading ? "Indexando los cuatro archivos…" : "Pregunta con tus propias palabras…"}
            />
            <button onClick={() => ask()} disabled={!question.trim() || loading}>Enviar</button>
            <small>
              Admite variantes como “cuántas OT”, “órdenes por estado”, “quién tiene más horas”
              o “grafica el costo por técnico”.
            </small>
          </div>
        </section>
      </div>
    </main>
  );
}
