import { NextResponse } from "next/server";
import { buildCatalog, loadTables } from "../../../lib/data";
import { createQueryPlan } from "../../../lib/openai";
import { executePlan } from "../../../lib/query";

export const runtime = "nodejs";
export const maxDuration = 30;

export async function POST(request: Request) {
  try {
    const body = await request.json() as { question?: unknown };
    const question = typeof body.question === "string" ? body.question.trim() : "";
    if (question.length < 3) return NextResponse.json({ error: "Escribe una pregunta más completa." }, { status: 400 });
    if (question.length > 500) return NextResponse.json({ error: "La pregunta es demasiado extensa." }, { status: 400 });
    const tables = loadTables();
    const plan = await createQueryPlan(question, buildCatalog(tables));
    return NextResponse.json(executePlan(plan, tables));
  } catch (error) {
    const message = error instanceof Error ? error.message : "No fue posible procesar la pregunta.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
