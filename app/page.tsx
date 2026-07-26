"use client";

import { FormEvent, useState } from "react";

type Message = {
  role: "user" | "assistant";
  text: string;
  interpretation?: string;
  sources?: string[];
  table?: Array<Record<string, unknown>>;
  warning?: string;
};

const examples = [
  "¿Cuántas órdenes de trabajo tiene el taller?",
  "¿Cuántas personas de control de calidad hay?",
  "Distribución de las OT por estado",
  "Promedio de horas productivas por técnico"
];

export default function Home() {
  const [question, setQuestion] = useState("");
  const [loading, setLoading] = useState(false);
  const [messages, setMessages] = useState<Message[]>([{
    role: "assistant",
    text: "Listo. El modelo de IA interpretará tu pregunta y el motor de datos verificará la respuesta en los cuatro archivos Excel."
  }]);

  async function ask(event?: FormEvent) {
    event?.preventDefault();
    const value = question.trim();
    if (!value || loading) return;
    setMessages(current => [...current, { role: "user", text: value }]);
    setQuestion("");
    setLoading(true);
    try {
      const response = await fetch("/api/ask", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ question: value })
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "No fue posible responder.");
      setMessages(current => [...current, {
        role: "assistant",
        text: data.answer,
        interpretation: data.interpretation,
        sources: data.sources,
        table: data.table,
        warning: data.warning
      }]);
    } catch (error) {
      setMessages(current => [...current, {
        role: "assistant",
        text: error instanceof Error ? error.message : "Ocurrió un error inesperado."
      }]);
    } finally {
      setLoading(false);
    }
  }

  return (
    <main>
      <header className="topbar">
        <div className="brand"><span>EQ</span><strong>ExcelQ</strong><small>Asistente con IA</small></div>
        <div className="status"><i /> 4 libros Excel conectados</div>
      </header>

      <section className="hero">
        <div>
          <p className="eyebrow">UNIVERSIDAD · DATOS · INTELIGENCIA ARTIFICIAL</p>
          <h1>Universidad de las Fuerzas Armadas ESPE Latacunga<br /><em>Mantenimiento Automotriz</em></h1>
          <p className="intro">La IA interpreta preguntas en lenguaje natural; el motor consulta, calcula y contrasta la información de todos los archivos.</p>
        </div>
        <aside>
          <strong>Respuesta verificable</strong>
          <p>Las cantidades se calculan desde los Excel. Si dos archivos no coinciden, el asistente muestra la diferencia.</p>
        </aside>
      </section>

      <section className="workspace">
        <div className="chatHeader">
          <div><span className="bot">IA</span><strong>Asistente de análisis integral</strong></div>
          <span className="ready">{loading ? "Analizando" : "Listo"}</span>
        </div>
        <div className="messages">
          {messages.map((message, index) => (
            <article key={index} className={message.role}>
              <span className="avatar">{message.role === "user" ? "TÚ" : "IA"}</span>
              <div className="bubble">
                <p>{message.text}</p>
                {message.interpretation && <p className="interpretation"><b>Interpretación:</b> {message.interpretation}</p>}
                {message.warning && <p className="warning">{message.warning}</p>}
                {message.table?.length ? (
                  <div className="tableWrap"><table>
                    <thead><tr>{Object.keys(message.table[0]).map(key => <th key={key}>{key}</th>)}</tr></thead>
                    <tbody>{message.table.map((row, rowIndex) => <tr key={rowIndex}>
                      {Object.keys(message.table![0]).map(key => <td key={key}>{String(row[key] ?? "")}</td>)}
                    </tr>)}</tbody>
                  </table></div>
                ) : null}
                {message.sources?.length ? <p className="sources"><b>Fuentes:</b> {[...new Set(message.sources)].join(" · ")}</p> : null}
              </div>
            </article>
          ))}
          {loading && <article className="assistant"><span className="avatar">IA</span><div className="bubble thinking"><i /><i /><i /></div></article>}
        </div>

        <form onSubmit={ask}>
          <input aria-label="Pregunta para el asistente" value={question} onChange={event => setQuestion(event.target.value)}
            placeholder="Pregunta sobre cualquiera de los cuatro Excel…" disabled={loading} />
          <button disabled={loading || !question.trim()}>{loading ? "Procesando…" : "Enviar"}</button>
        </form>
        <div className="examples">{examples.map(example => <button key={example} onClick={() => setQuestion(example)}>{example}</button>)}</div>
      </section>
    </main>
  );
}
