import type { QueryPlan } from "./types";

const schema = {
  type: "object",
  additionalProperties: false,
  required: ["operation", "subject", "target_column", "group_by", "preferred_sheet", "filters", "interpretation"],
  properties: {
    operation: { type: "string", enum: ["count_unique", "count_rows", "list", "sum", "average", "min", "max", "group_count", "lookup", "unsupported"] },
    subject: { type: "string" },
    target_column: { type: ["string", "null"] },
    group_by: { type: ["string", "null"] },
    preferred_sheet: { type: ["string", "null"] },
    filters: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["column", "operator", "value"],
        properties: {
          column: { type: "string" },
          operator: { type: "string", enum: ["equals", "contains"] },
          value: { type: "string" }
        }
      }
    },
    interpretation: { type: "string" }
  }
};

function extractText(response: Record<string, unknown>) {
  if (typeof response.output_text === "string") return response.output_text;
  const output = Array.isArray(response.output) ? response.output : [];
  for (const item of output as Array<Record<string, unknown>>) {
    const content = Array.isArray(item.content) ? item.content : [];
    for (const block of content as Array<Record<string, unknown>>) {
      if (block.type === "output_text" && typeof block.text === "string") return block.text;
    }
  }
  throw new Error("El modelo no devolvió un plan de consulta.");
}

export async function createQueryPlan(question: string, catalog: unknown): Promise<QueryPlan> {
  const key = process.env.OPENAI_API_KEY;
  if (!key) throw new Error("OPENAI_API_KEY no está configurada en Vercel.");
  const model = process.env.OPENAI_MODEL || "gpt-5.6-sol";
  const instructions = `Eres el planificador de consultas de un taller automotriz de la Universidad de las Fuerzas Armadas ESPE.
Interpreta preguntas en español y crea un plan usando EXCLUSIVAMENTE los nombres de columnas y valores del catálogo.
No calcules ni inventes respuestas. El servidor hará los cálculos exactos.
Reglas:
- OT, órdenes de trabajo y vehículos registrados normalmente se cuentan por la columna OT.
- Para personal usa la hoja Personal, Nombre y Cargo.
- "cuántos técnicos/lavadores/asesores/control de calidad" significa contar Nombre filtrando Cargo.
- Estados como Parada, Planificada, Proyectada, Prevista y En desarrollo son filtros de Estado.
- Si se pide "por técnico/estado/tipo", usa group_count o una agregación con group_by.
- Si el dato no está en el catálogo, usa unsupported.
- preferred_sheet solo se usa cuando la pregunta nombra o implica claramente una hoja.
CATÁLOGO:
${JSON.stringify(catalog)}`;

  const apiResponse = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${key}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model,
      instructions,
      input: question,
      text: {
        format: {
          type: "json_schema",
          name: "query_plan",
          strict: true,
          schema
        }
      }
    })
  });
  const payload = await apiResponse.json() as Record<string, unknown>;
  if (!apiResponse.ok) {
    const detail = typeof payload.error === "object" && payload.error && "message" in payload.error
      ? String((payload.error as { message: unknown }).message)
      : `Error ${apiResponse.status}`;
    throw new Error(`OpenAI: ${detail}`);
  }
  return JSON.parse(extractText(payload)) as QueryPlan;
}
