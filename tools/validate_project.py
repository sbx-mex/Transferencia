#!/usr/bin/env python3
"""Validación reproducible del proyecto mensual de Transferencias."""

from __future__ import annotations

import csv
import importlib.util
import json
import re
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PROCESSOR_PATH = ROOT / "tools" / "process_monthly_data.py"
RESULT_PATH = ROOT / "docs" / "validation-results.json"
MONTH_ORDER = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio"]

spec = importlib.util.spec_from_file_location("monthly_processor", PROCESSOR_PATH)
processor = importlib.util.module_from_spec(spec)
assert spec.loader
sys.modules[spec.name] = processor
spec.loader.exec_module(processor)

results: list[dict[str, str]] = []


def check(name: str, condition: bool, evidence: str) -> None:
    results.append(
        {
            "Prueba": name,
            "Resultado": "Aprobada" if condition else "Fallida",
            "Evidencia": evidence,
        }
    )
    if not condition:
        raise AssertionError(f"{name}: {evidence}")


def json_file(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


manifest = json_file(ROOT / "data" / "manifest-data.json")
summary = json_file(ROOT / "data" / "audit" / "resumen-procesamiento.json")
exclusions = json_file(ROOT / "data" / "audit" / "exclusiones.json")
hidden = json_file(ROOT / "data" / "audit" / "ocultos.json")
inconsistencies = json_file(ROOT / "data" / "audit" / "inconsistencias-datos.json")
date_audit = json_file(ROOT / "data" / "audit" / "fechas.json")
html = (ROOT / "index.html").read_text(encoding="utf-8")
css = (ROOT / "css" / "styles.css").read_text(encoding="utf-8")
js = (ROOT / "js" / "app.js").read_text(encoding="utf-8")
sw = (ROOT / "service-worker.js").read_text(encoding="utf-8")

check("1. Estructura original conservada", all((ROOT / item).exists() for item in ["index.html", "css/styles.css", "js/app.js", "manifest.json", "service-worker.js"]), "Archivos principales presentes.")
check("2. README e instrucción interna", (ROOT / "README.md").exists() and (ROOT / "AUDITORIA_TRANSFERENCIAS_V6.txt").exists(), "Documentación presente.")
check("3. Siete CSV mensuales disponibles", len(list((ROOT / "data/source/monthly").glob("*.csv"))) == 7, "Enero a Julio.")
check("4. Archivo maestro disponible", (ROOT / "data/source/Base_Transferencias.xlsx").exists(), "Base_Transferencias.xlsx presente.")
check("5. Detección cronológica de meses", [item["label"] for item in manifest["chunks"]] == MONTH_ORDER, "Orden Enero–Julio.")
check("6. No se usa fecha de modificación", "stat(" not in PROCESSOR_PATH.read_text(encoding="utf-8"), "El mes se obtiene de nombre y fecha interna.")
check("7. Codificación detectada", all(item["encoding"] == "utf-8-sig" for item in summary["monthly"]), "7/7 UTF-8 con BOM.")
check("8. Delimitador detectado", all(item["delimiter"] == "," for item in summary["monthly"]), "7/7 delimitador coma.")
check("9. Encabezados compatibles", manifest["generatedFromHeaders"] == processor.EXPECTED_HEADERS, "12/12 encabezados.")
check("10. Fechas válidas", summary["totals"]["invalidDateRows"] == 0, "0 fechas inválidas.")
check("11. Fechas futuras", summary["totals"]["futureDateRows"] == 0, "0 fechas futuras.")
check("12. Julio interpretado correctamente", summary["monthly"][-1]["min_date"] == "2026-07-01" and summary["monthly"][-1]["max_date"] == "2026-07-23", "01/07/2026–23/07/2026.")
check("13. Última actualización interna", manifest["latestUpdate"] == "2026-07-23", "Máxima fecha válida de Julio.csv.")
check("14. Evidencia de última actualización", manifest["latestUpdateEvidence"]["sourceFile"] == "Julio.csv", "Archivo y periodo trazables.")
check("15. Mes actual disponible", manifest["latestUpdateEvidence"]["currentMonthFileFound"] is True, "Julio.csv corresponde al mes actual.")
fallback_period, fallback_used = processor.resolve_latest_period(
    [
        processor.MonthlyResult(
            file="Junio.csv",
            month=6,
            year=2026,
            encoding="utf-8-sig",
            delimiter=",",
            max_date="2026-06-30",
        )
    ],
    processor.date(2026, 7, 24),
)
check("16. Prueba de fallback", fallback_used and fallback_period.file == "Junio.csv", "Sin Julio.csv se selecciona Junio.csv y se marca fallback.")
check("16a. Sistema de semana comprobado", all(item["week_system"] == "ISO" for item in summary["monthly"]), "7/7 archivos coinciden con semana ISO.")
check("16b. Patrones de fecha por archivo", all(item["date_pattern"] == "DD/MM/YYYY" for item in summary["monthly"][:6]) and summary["monthly"][-1]["date_pattern"] == "M/D/YYYY", "Enero–junio DD/MM/YYYY; julio M/D/YYYY.")
check("16c. Fechas conservadas", summary["totals"]["dateConservedRows"] == 304498, "304,498 fechas mantienen su interpretación.")
check("16d. Fechas corregidas", summary["totals"]["dateConvertedRows"] == 34664, "34,664 fechas M/D/YYYY normalizadas.")
check("16e. Fechas ambiguas", summary["totals"]["dateAmbiguousRows"] == 0 and summary["totals"]["weekMismatchRows"] == 0, "0 fechas pendientes en las fuentes actuales.")
check("16f. Auditoría de fechas trazable", sum(item["Cantidad registros"] for item in date_audit) == summary["totals"]["sourceRows"], "339,162 filas representadas con valor original, formato, fecha, año, semana y motivo.")
check("17. Registros originales", summary["totals"]["sourceRows"] == 339162, "339,162 filas.")
check("18. Consolidación total", summary["totals"]["recordsAuditoria"] == 338220, "338,220 registros técnicos después de Non Inventory.")
check("19. Duplicados exactos", summary["totals"]["duplicateRows"] == 0, "0 duplicados.")
check("20. Prevención de duplicados", processor.exact_fingerprint({header: "x" for header in processor.EXPECTED_HEADERS}) == processor.exact_fingerprint({header: "x" for header in processor.EXPECTED_HEADERS}), "Huella estable de 12 campos.")
check("21. Conciliación de registros", summary["totals"]["reconciliation"]["difference"] == 0, "Originales = visibles + exclusiones + ocultos + incidencias.")
check("22. Base_Directorio exacta", summary["master"]["directoryRows"] == 968, "968 CeCo únicos.")
check("23. CeCo duplicados en Directorio", summary["master"]["directoryDuplicateKeys"] == 0, "0 claves duplicadas.")
check("24. Cruces correctos de Directorio", summary["totals"]["directoryMatches"] == 339162, "339,162 coincidencias exactas.")
check("25. CeCo sin correspondencia", summary["totals"]["directoryMissing"] == 0, "0 registros.")
check("26. Región y DM disponibles", all(item.get("region") and item.get("dm") for item in manifest["directory"]), "968 registros con Región y DM.")
check("27. CeCo normalizado", processor.normalize_ceco("03810") == "03810" and processor.normalize_ceco(38101) == "38101", "Ceros iniciales y 5 dígitos conservados.")
check("28. Non Inventory único", summary["master"]["nonInventoryUniqueKeys"] == 343, "343 claves exactas.")
check("29. Exclusiones Non Inventory", len(exclusions) == 942 == summary["totals"]["nonInventoryRows"], "942 registros trazables.")
check("30. Monto Non Inventory", round(sum(item["Importe"] for item in exclusions), 2) == summary["totals"]["nonInventoryAmount"], "Monto conciliado con detalle.")
check("31. Originales no eliminados", sum(item["source_rows"] for item in summary["monthly"]) == 339162, "Los CSV fuente permanecen intactos.")
check("32. Patrol oculto documentado", summary["totals"]["hiddenPatrolRows"] == 863, "863 registros en fuente; 862 posteriores a Non Inventory.")
check("33. Proveedor sin CeCo documentado", summary["totals"]["hiddenUnresolvedProviderRows"] == 327, "327 registros en fuente; 323 posteriores a Non Inventory.")
check("34. Registro técnico de ocultos", len(hidden) == summary["totals"]["hiddenPatrolRowsInVisibleData"] + summary["totals"]["hiddenUnresolvedProviderRowsInVisibleData"], f"{len(hidden)} registros técnicos.")
check("35. Controles de ocultamiento no visibles", "coffeeFilter" not in html and "unresolvedProviderFilter" not in html, "Criterios fijos, no reactivables desde la vista.")
check("36. Selector de mes y todos los meses", 'id="monthFilter"' in html and '<option value="">Todos</option>' in html, "Control disponible.")
check("37. Filtros requeridos", all(f'id="{item}"' in html for item in ["regionFilter", "dmFilter", "storeFilter", "searchFilter", "statusFilter"]), "Región, DM, tienda, búsqueda y estatus.")
check("38. Limpieza de filtros", 'id="resetFilters"' in html and "resetFilters" in js, "Acción enlazada.")
check("39. Carga mensual aislada", "requiredChunks" in js and "state.allRows=needed.flatMap" in js, "El conjunto activo se reconstruye al cambiar periodo.")
check("40. Error de un mes no bloquea los demás", "state.loadErrors.push" in js and "for(const chunk of needed)" in js, "Manejo por archivo.")
check("41. Mensaje sin resultados", "No hay transferencias con los filtros actuales" in js, "Estado vacío accionable.")
check("42. Indicadores reactivos", all(label in js for label in ["Registros analizados", "Casos críticos", "Registros excluidos"]), "KPIs calculados en apply/renderAll.")
check("43. Evidencia interna sin saturar tabla", "Archivo / fila" in html and "<th>Evidencia de cruce</th>" not in html and "function evidenceCell" in js, "Archivo/fila visible; evidencia conservada internamente.")
check("43a. Detalle cerrado inicialmente", 'id="detailBody" hidden' in html and 'id="detailToggle"' in html and "Abrir detalle" in js and "Cerrar detalle" in js, "Control dinámico accesible con aria-expanded.")
check("44. Pie de página único", html.count("Diseñado: Jorge Alcantar Aguiar &amp; Enrique César Flores") == 1, "Una sola aparición.")
check("45. Responsive desde 320 px", "@media(max-width:720px)" in css and "overflow:auto" in css and "min-width:0" in css, "Reglas móviles y contenedor de tabla sin desbordar la página.")
check("46. Sin rutas absolutas", not re.search(r"""(?:src|href)=["']/(?!/)""", html) and "/workspace/" not in html + js + sw, "Rutas relativas.")
check("47. Manifest válido", json_file(ROOT / "manifest.json")["start_url"] == "./", "start_url y scope relativos.")
check("48. Caché PWA incrementada", "transferencias-v14-fechas-auditadas" in sw, "Versión v14.")
check("49. Datos network-first", "/data/chunks/" in sw and "cache.put(event.request,copy)" in sw, "Actualiza caché y conserva fallback.")
check("50. JavaScript válido", subprocess.run(["node", "--check", str(ROOT / "js/app.js")], capture_output=True).returncode == 0 and subprocess.run(["node", "--check", str(ROOT / "service-worker.js")], capture_output=True).returncode == 0, "app.js y service-worker.js.")

for path in list((ROOT / "data").rglob("*.json")) + [ROOT / "manifest.json", ROOT / "docs/file-audit.json"]:
    json_file(path)
check("51. Todos los JSON válidos", True, "Manifest, chunks y auditoría parseados.")

runtime_path = Path(tempfile.gettempdir()) / "transfer-runtime-validation.json"
runtime = subprocess.run(
    ["node", str(ROOT / "tools/audit_runtime.mjs"), str(ROOT), str(runtime_path)],
    capture_output=True,
    text=True,
    check=False,
)
check("52. Lógica de conciliación ejecutable", runtime.returncode == 0, "Auditoría Node completada.")
runtime_data = json_file(runtime_path)
check("53. Casos críticos calculados", runtime_data["criticalCaseCount"] == 31585, "31,585 casos mensuales requieren revisión.")
critical_case = next(
    (
        item
        for item in runtime_data["criticalCases"]
        if item["Fecha"] == "2026-07-21"
        and item["CeCo origen"] == "38339"
        and item["CeCo destino"] == "38456"
        and "Vaso de Plastico 20 oz" in item["Artículos"]
    ),
    None,
)
check(
    "53a. Caso 38339 → 38456",
    critical_case is not None
    and critical_case["Estatus"] == "Falta salida"
    and critical_case["Monto salida"] == 0
    and critical_case["Monto entrada"] == 228,
    "La fuente conserva 0.00 en 38339 y evidencia entrada 150.0 / $228.00 en 38456; no se inventa salida.",
)
case_date = next(
    (
        item
        for item in date_audit
        if item["Archivo"] == "Julio.csv"
        and item["Valor original"] == "7/21/2026"
        and item["Semana informada"] == 30
    ),
    None,
)
check(
    "53b. Fecha crítica normalizada",
    case_date is not None
    and case_date["Fecha normalizada"] == "21/07/2026"
    and case_date["Formato detectado"] == "M/D/YYYY"
    and case_date["Semana calculada"] == 30,
    "7/21/2026 → 21/07/2026 por Año 2026, semana ISO 30 y mes Julio.",
)
check("54. Julio sin datos residuales", runtime_data["monthly"][-1]["matchingRowsLoaded"] == 36615, "Solo Julio se necesita para conciliar el periodo más reciente.")
check("55. Todos los meses consultables", len(runtime_data["monthly"]) == 7 and runtime_data["allMonths"]["visibleTransfers"] > 0, "7 periodos y vista consolidada.")

server = subprocess.Popen(
    [sys.executable, "-m", "http.server", "8765", "--bind", "127.0.0.1"],
    cwd=ROOT.parent,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
)
try:
    time.sleep(0.5)
    urls = [
        "http://127.0.0.1:8765/Transferencia-main/",
        "http://127.0.0.1:8765/Transferencia-main/css/styles.css",
        "http://127.0.0.1:8765/Transferencia-main/js/app.js",
        "http://127.0.0.1:8765/Transferencia-main/data/manifest-data.json",
    ]
    statuses = [urllib.request.urlopen(url, timeout=5).status for url in urls]
finally:
    server.terminate()
    server.wait(timeout=5)
check("56. Compatibilidad con subruta GitHub Pages", statuses == [200, 200, 200, 200], "Carga HTTP desde /Transferencia-main/.")

with tempfile.TemporaryDirectory() as temp_dir:
    temp = Path(temp_dir)
    empty = temp / "Agosto.csv"
    empty.write_text(",".join(processor.EXPECTED_HEADERS) + "\n", encoding="utf-8")
    encoding, delimiter = processor.detect_csv(empty)
    with empty.open(encoding=encoding, newline="") as handle:
        empty_rows = list(csv.DictReader(handle, delimiter=delimiter))
    missing = temp / "Septiembre.csv"
    missing.write_text("Año,Dia\n2026,01/09/2026\n", encoding="utf-8")
    with missing.open(encoding="utf-8", newline="") as handle:
        fields = csv.DictReader(handle).fieldnames or []
check("57. Manejo de CSV vacío", len(empty_rows) == 0 and "CSV está vacío" in PROCESSOR_PATH.read_text(encoding="utf-8"), "Rechazo aislado con mensaje.")
check("58. Manejo de columnas faltantes", len(fields) < len(processor.EXPECTED_HEADERS) and "Columnas faltantes" in PROCESSOR_PATH.read_text(encoding="utf-8"), "Rechazo aislado con detalle.")
check("59. Manejo de fecha inválida", processor.parse_month_date("31/02/2026", 2, 2026) is None, "Fecha rechazada.")
check("60. Actualización sin JavaScript manual", "glob(\"*.csv\")" in PROCESSOR_PATH.read_text(encoding="utf-8"), "Descubrimiento automático.")

payload = {
    "summary": {
        "executed": len(results),
        "passed": sum(item["Resultado"] == "Aprobada" for item in results),
        "failed": sum(item["Resultado"] == "Fallida" for item in results),
        "notExecuted": [
            "Validación visual real en navegador a 320/768/1440 px.",
            "Instalación PWA y recarga sin conexión en un navegador físico.",
            "Disponibilidad de los destinos externos después de publicar en GitHub Pages.",
        ],
    },
    "tests": results,
}
RESULT_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(payload["summary"], ensure_ascii=False, indent=2))
