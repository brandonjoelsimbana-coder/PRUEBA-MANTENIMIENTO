import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import ts from "typescript";

const require = createRequire(import.meta.url);
process.env.NODE_PATH = path.join(process.cwd(), "node_modules");
require("node:module").Module._initPaths();

const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), "espe-engine-"));
const compiledPath = path.join(temporaryDirectory, "excel-engine.cjs");

try {
  const source = fs.readFileSync(path.join(process.cwd(), "lib", "excel-engine.ts"), "utf8");
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true,
    },
  });
  fs.writeFileSync(compiledPath, compiled.outputText);
  const engine = require(compiledPath);
  const files = [
    ["Entregable 1 · Base de datos", "Planning_Carga_Trabajo_Entregable_1.xlsx"],
    ["Planning generado por IA", "Planning_IA.xlsx"],
    ["Entregable 3 · Dashboard", "Entregable_3_Dashboard_Indicadores.xlsx"],
    ["Entregable 2 · Planning", "Entregable_2_Planning_Carga_Trabajo.xlsx"],
  ];
  const sources = files.flatMap(([name, file]) => {
    const bytes = fs.readFileSync(path.join(process.cwd(), "public", "data", file));
    const data = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    return engine.parseWorkbook(data, name, file);
  });
  const cases = [
    ["¿Cuántos vehículos de tipo camioneta están registrados en el taller?", "Hay 33 vehículos de tipo camioneta"],
    ["¿Cuántas camionetas hay?", "Hay 33 vehículos de tipo camioneta"],
    ["Número de pickups", "Hay 33 vehículos de tipo camioneta"],
    ["¿Cuántos SUV hay?", "Hay 103 vehículos de tipo SUV"],
    ["Cantidad de automóviles", "Hay 83 vehículos de tipo automóvil"],
    ["¿Cuántos monovolúmenes existen?", "Hay 11 vehículos de tipo monovolumen"],
    ["¿Cuántas camionetas están paradas?", "Hay 1 vehículo de tipo camioneta"],
    ["¿Qué tipos de vehículos hay?", "4 tipos de vehículos"],
    ["Camionetas por estado", "camioneta se distribuyen por Estado"],
    ["¿Cuál es el número total de órdenes?", "tiene 230 órdenes de trabajo"],
    ["¿Cuántas órdenes están Planificada?", "73 órdenes de trabajo"],
    ["¿Cuántas órdenes de mantenimiento de 20 mil km hay?", "Hay 42 órdenes"],
    ["¿Qué tipos de mantenimiento hay?", "6 tipos de trabajo"],
    ["¿Cuántas órdenes con horas extra hay?", "Hay 46 órdenes con horas extra"],
    ["¿Cuántos conflictos de técnico existen?", "Hay 182 conflictos de técnico"],
    ["¿Cuántas entregas incumplidas hay?", "Hay 0 entregas incumplidas"],
    ["¿Cuántos técnicos tiene el taller?", "Hay 6 técnicos"],
    ["¿Cuántos lavadores hay?", "Hay 1 lavador"],
    ["¿Cuántas personas de control de calidad hay?", "Hay 1 persona"],
    ["¿Cuántas horas trabajó al mes cada técnico en el taller?", "Marie Dubois: 133,4 h"],
    ["Horas laboradas por técnico", "Willem Janssens: 131,22 h"],
    ["Total de horas extra por técnico", "Pierre Dubois: 20,07 h"],
    ["Total de horas normales por técnico", "Sophie Janssens: 100,53 h"],
    ["Total de horas requeridas por técnico", "Willem Janssens: 117,37 h"],
    ["¿Qué técnico trabajó más horas?", "Jan Peeters tiene el mayor acumulado de Horas productivas: 136,11"],
    ["¿Qué técnico trabajó menos horas?", "Sophie Janssens tiene el menor acumulado de Horas productivas: 123,94"],
    ["Total de costo control de calidad por estado", "total de Costo control calidad por Estado"],
    ["¿Cuántas órdenes están detenidas?", "4 órdenes"],
    ["¿Cuántas grúas hay?", "No encontré"],
  ];

  let failures = 0;
  for (const [question, expected] of cases) {
    const answer = engine.attachChart(
      engine.answerAcross(question, sources),
      question,
    );
    if (!answer.text.includes(expected)) {
      failures += 1;
      console.error(`FALLO: ${question}\n${answer.text}`);
    }
  }
  const monthly = engine.answerAcross(
    "¿Cuántas horas trabajó al mes cada técnico en el taller?",
    sources,
  );
  if (
    !monthly.text.includes("no cubren el mes completo") ||
    monthly.table?.length !== 6
  ) {
    failures += 1;
    console.error("FALLO: la respuesta mensual debe aclarar el período parcial y listar 6 técnicos.");
  }
  const catalog = engine.buildQuestionCatalog(sources);
  const catalogCount = catalog.reduce(
    (total, category) => total + category.questions.length,
    0,
  );
  if (catalog.length < 10 || catalogCount < 250) {
    failures += 1;
    console.error(`FALLO: catálogo insuficiente (${catalog.length} temas, ${catalogCount} preguntas).`);
  }
  if (failures) process.exitCode = 1;
  else console.log(`OK: ${cases.length} preguntas verificadas y ${catalogCount} variantes en ${sources.length} hojas.`);
} finally {
  fs.rmSync(temporaryDirectory, { recursive: true, force: true });
}
