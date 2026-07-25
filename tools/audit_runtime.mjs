#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { performance } from "node:perf_hooks";

const projectRoot = path.resolve(process.argv[2] || path.join(import.meta.dirname, ".."));
const outputPath = process.argv[3] ? path.resolve(process.argv[3]) : "";
const manifest = JSON.parse(fs.readFileSync(path.join(projectRoot, "data/manifest-data.json"), "utf8"));
let source = fs.readFileSync(path.join(projectRoot, "js/app.js"), "utf8");
if (!source.includes("__test:{")) {
  source = source.replace(
    "return{state,apply};",
    "return{state,apply,__test:{buildAllTransfers,applyTransferFilters,computeMetrics,rowMatchesTime,isCoffeeProvider,isUnresolvedProvider}};",
  );
}
source += "\nglobalThis.__auditApp=app;";

const sandbox = {
  console,
  Intl,
  Date,
  Map,
  Set,
  Math,
  Number,
  String,
  Array,
  Object,
  RegExp,
  JSON,
  setTimeout,
  clearTimeout,
  document: {
    addEventListener() {},
    getElementById() {
      return null;
    },
    querySelectorAll() {
      return [];
    },
  },
  navigator: {},
  window: {},
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: "app.js" });
const app = sandbox.__auditApp;
const { state } = app;
const test = app.__test;
state.manifest = manifest;
for (const item of manifest.directory || []) state.dirByCeco.set(String(item.ceco), item);
for (const [ceco, store] of Object.entries(manifest.cecoStores || {})) {
  state.storeByCeco.set(String(ceco), store);
}

const rowsByChunk = new Map();
for (const chunk of manifest.chunks || []) {
  const payload = JSON.parse(fs.readFileSync(path.join(projectRoot, chunk.path), "utf8"));
  rowsByChunk.set(chunk.id, payload.rows || []);
}

function rowsForPeriod(year, month) {
  return (manifest.chunks || [])
    .filter((chunk) => chunk.year > year || (chunk.year === year && chunk.month >= month))
    .flatMap((chunk) => rowsByChunk.get(chunk.id) || []);
}

function simplifyMetrics(metrics) {
  return Object.fromEntries(
    Object.entries(metrics).map(([key, value]) => [
      key,
      typeof value === "number" ? Math.round(value * 100) / 100 : value,
    ]),
  );
}

const monthly = [];
const detailCases = [];
const totalStarted = performance.now();
for (const monthItem of manifest.months || []) {
  const year = Number(monthItem.year || manifest.chunks.find((item) => item.month === monthItem.value)?.year);
  const month = Number(monthItem.value);
  state.year = String(year);
  state.month = String(month);
  state.week = "";
  state.region = "";
  state.dm = "";
  state.store = "";
  state.ingredient = "";
  state.provider = "";
  state.status = "";
  state.direction = "";
  state.search = "";
  state.hideCoffee = true;
  state.hideUnresolvedProvider = true;
  const rows = rowsForPeriod(year, month);
  const started = performance.now();
  const allTransfers = test.buildAllTransfers(rows);
  const visibleTransfers = test.applyTransferFilters(allTransfers);
  const elapsed = performance.now() - started;
  const metrics = test.computeMetrics(visibleTransfers);
  monthly.push({
    year,
    month,
    label: monthItem.label,
    sourceFile: monthItem.sourceFile,
    matchingRowsLoaded: rows.length,
    transfersBeforeOperationalHiding: allTransfers.length,
    visibleTransfers: visibleTransfers.length,
    elapsedMilliseconds: Math.round(elapsed * 100) / 100,
    metrics: simplifyMetrics(metrics),
  });
  for (const transfer of visibleTransfers) {
    if (test.computeMetrics([transfer]).pendientes === 0) continue;
    const first = transfer.products[0]?.salida || transfer.products[0]?.ingreso || [];
    const files = [
      ...new Set(
        transfer.products.flatMap((product) =>
          [product.salida, product.ingreso]
            .filter(Boolean)
            .map((row) => {
              const raw = row[manifest.idx?.fileOrigin ?? 14];
              return manifest.dicts.archivo?.[raw] || raw;
            })
            .filter(Boolean),
        ),
      ),
    ];
    const sourceRows = [
      ...new Set(
        transfer.products.flatMap((product) =>
          [product.salida, product.ingreso]
            .filter(Boolean)
            .map((row) => row[manifest.idx?.sourceRow ?? 15])
            .filter(Boolean),
        ),
      ),
    ];
    const ingredients = [
      ...new Set(
        transfer.products
          .map((product) => {
            const row = product.salida || product.ingreso;
            return manifest.dicts.ingrediente[row?.[manifest.idx?.ingredient ?? 5]] || "";
          })
          .filter(Boolean),
      ),
    ];
    detailCases.push({
      Mes: monthItem.label,
      Fecha: transfer.date,
      "CeCo origen": transfer.originCeco,
      "Tienda origen": transfer.originStore,
      "CeCo destino": transfer.destinationCeco,
      "Tienda destino": transfer.destinationStore,
      Región: transfer.originRegion || transfer.destinationRegion,
      DM: transfer.originDm || transfer.destinationDm,
      "Documento / folio": "",
      Artículos: ingredients.join(" | "),
      "Cantidad de artículos": transfer.productCount,
      "Monto salida": Math.round(transfer.montoSalida * 100) / 100,
      "Monto entrada": Math.round(transfer.montoIngreso * 100) / 100,
      Diferencia: Math.round(transfer.diff * 100) / 100,
      "Tipo movimiento": transfer.type,
      Estatus: transfer.status,
      Causa: transfer.status,
      "Archivo origen": files.join(" | "),
      "Filas origen": sourceRows.join(" | "),
      "Evidencia cruce": `${transfer.originCeco || "Sin CeCo"} → Base_Directorio → ${transfer.originRegion || "Sin Región"} → ${transfer.originDm || "Sin DM"}`,
      "Resultado auditoría": "Requiere revisión manual",
      Riesgo: ["Sin ingreso", "Falta salida", "Diferencia de cantidad", "Diferencia de costo", "Parcial"].includes(transfer.status)
        ? "Alto"
        : "Medio",
    });
  }
}

state.year = "";
state.month = "";
const allRows = [...rowsByChunk.values()].flat();
const allStarted = performance.now();
const allTransfers = test.buildAllTransfers(allRows);
const visibleTransfers = test.applyTransferFilters(allTransfers);
const allElapsed = performance.now() - allStarted;
const output = {
  generatedAt: new Date().toISOString(),
  projectVersion: manifest.version,
  rowsLoaded: allRows.length,
  allMonths: {
    transfersBeforeOperationalHiding: allTransfers.length,
    visibleTransfers: visibleTransfers.length,
    elapsedMilliseconds: Math.round(allElapsed * 100) / 100,
    metrics: simplifyMetrics(test.computeMetrics(visibleTransfers)),
  },
  monthly,
  criticalCaseCount: detailCases.length,
  totalElapsedMilliseconds: Math.round((performance.now() - totalStarted) * 100) / 100,
  criticalCases: outputPath ? detailCases : undefined,
};
if (outputPath) fs.writeFileSync(outputPath, JSON.stringify(output));
process.stdout.write(`${JSON.stringify(outputPath ? { ...output, criticalCases: undefined } : output, null, 2)}\n`);
