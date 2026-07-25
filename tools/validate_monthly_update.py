#!/usr/bin/env python3
"""Valida una regeneración mensual sin depender de cifras fijas de un periodo."""

from __future__ import annotations

import json
import subprocess
import sys
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MONTHLY_ROOT = ROOT / "data" / "source" / "monthly"
MANIFEST_PATH = ROOT / "data" / "manifest-data.json"
SUMMARY_PATH = ROOT / "data" / "audit" / "resumen-procesamiento.json"
FILE_AUDIT_PATH = ROOT / "docs" / "file-audit.json"


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def main() -> None:
    csv_files = sorted(path.name for path in MONTHLY_ROOT.glob("*.csv"))
    require(csv_files, "No hay archivos CSV en data/source/monthly.")

    manifest = load_json(MANIFEST_PATH)
    summary = load_json(SUMMARY_PATH)
    file_audit = load_json(FILE_AUDIT_PATH)

    monthly = summary.get("monthly", [])
    processed = [item for item in monthly if item.get("status") == "Procesado"]
    rejected = [item for item in monthly if item.get("status") != "Procesado"]

    require(len(monthly) == len(csv_files), "No se auditaron todos los CSV disponibles.")
    require(not rejected, f"Hay CSV rechazados: {[item.get('file') for item in rejected]}")
    require(
        {item.get("file") for item in processed} == set(csv_files),
        "Los CSV procesados no coinciden con los archivos fuente.",
    )
    require(
        summary["totals"]["processedFiles"] == len(csv_files),
        "El total de archivos procesados es incorrecto.",
    )
    require(summary["totals"]["rejectedFiles"] == 0, "Existen archivos rechazados.")
    require(summary["totals"]["invalidDateRows"] == 0, "Existen fechas inválidas.")
    require(summary["totals"]["dateAmbiguousRows"] == 0, "Existen fechas ambiguas.")
    require(summary["totals"]["weekMismatchRows"] == 0, "Existen fechas que no coinciden con Año/Semana.")
    require(summary["totals"]["directoryMissing"] == 0, "Existen CeCo sin correspondencia en el Directorio.")
    require(
        summary["totals"]["reconciliation"]["difference"] == 0,
        "La conciliación de registros presenta diferencias.",
    )
    require(
        summary["totals"]["sourceRows"] == sum(item["source_rows"] for item in processed),
        "El total de filas fuente no coincide con la suma mensual.",
    )

    chunks = manifest.get("chunks", [])
    require(len(chunks) == len(csv_files), "No se generó un bloque de datos por cada CSV.")
    require(
        [item["month"] for item in chunks] == sorted(item["month"] for item in chunks),
        "Los periodos no quedaron en orden cronológico.",
    )
    require(
        {item["sourceFile"] for item in chunks} == set(csv_files),
        "Los bloques generados no corresponden a todos los CSV.",
    )
    for chunk in chunks:
        path = ROOT / chunk["path"]
        require(path.exists(), f"No existe el bloque generado: {chunk['path']}")
        payload = load_json(path)
        require(payload["sourceFile"] == chunk["sourceFile"], f"Fuente incorrecta en {chunk['path']}.")
        require(len(payload["rows"]) == chunk["records"], f"Conteo incorrecto en {chunk['path']}.")

    latest = manifest.get("latestUpdate", "")
    require(latest, "No se determinó la última actualización.")
    require(date.fromisoformat(latest) <= date.today(), "La última actualización está en el futuro.")
    require(
        latest == summary["latestUpdate"]["date"],
        "La última actualización no coincide entre manifiesto y auditoría.",
    )

    audited_paths = [item["path"] for item in file_audit.get("files", [])]
    require(
        not any(path == ".git" or path.startswith(".git/") for path in audited_paths),
        "La auditoría de archivos incluyó metadatos internos de Git.",
    )

    for path in list((ROOT / "data").rglob("*.json")) + [
        ROOT / "manifest.json",
        FILE_AUDIT_PATH,
    ]:
        load_json(path)

    for path in [ROOT / "js" / "app.js", ROOT / "service-worker.js"]:
        result = subprocess.run(
            ["node", "--check", str(path)],
            capture_output=True,
            text=True,
            check=False,
        )
        require(result.returncode == 0, result.stderr or f"JavaScript inválido: {path}")

    print(
        json.dumps(
            {
                "status": "ok",
                "csv": len(csv_files),
                "rows": summary["totals"]["sourceRows"],
                "latestUpdate": latest,
                "chunks": len(chunks),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise
