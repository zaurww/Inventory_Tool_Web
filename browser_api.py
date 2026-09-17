"""In-memory adapter used by the Pyodide browser pilot."""
from __future__ import annotations

from pathlib import Path
import time

import inventory as inv
from prepare_input import prepare_input_bytes
from migrate_input import migrate_input_bytes
from report_locale import localize_issues


def process_workbook(input_bytes, input_name="Inventory_Input.xlsx", language="az"):
    """Prepare, calculate, and serialize an XLSX workbook entirely in memory."""
    started = time.perf_counter()
    original = bytes(input_bytes)
    updated = original
    assigned_codes = 0
    migrated = False
    safe_name = Path(str(input_name)).name or "Inventory_Input.xlsx"
    try:
        if language == 'az':
            updated = migrate_input_bytes(original)
            migrated = True
        else:
            inv.read_input_bytes(original)
        updated, assigned_codes = prepare_input_bytes(updated)
        raw, digest = inv.read_input_bytes(updated)
        result = inv.calculate(raw)
        elapsed = time.perf_counter() - started
        warnings = [row for row in result["tables"]["Checks"] if row["Severity"] == "WARNING"]
        return {
            "ok": True,
            "assigned_codes": assigned_codes,
            "migrated": migrated,
            "warning_count": len(warnings),
            "elapsed": elapsed,
            "digest": digest,
            "source_counts": result["source_counts"],
            "input_bytes": updated,
            "report_bytes": inv.write_report_bytes(result, safe_name, digest, elapsed, language),
            "checks_bytes": None,
            "issues": [],
        }
    except inv.InvalidInput as exc:
        return {
            "ok": False,
            "assigned_codes": assigned_codes,
            "migrated": migrated,
            "warning_count": 0,
            "elapsed": time.perf_counter() - started,
            "digest": "",
            "source_counts": {},
            "input_bytes": updated,
            "report_bytes": None,
            "checks_bytes": inv.write_errors_bytes(exc.issues, language),
            "issues": localize_issues(exc.issues) if language == 'az' else exc.issues,
        }
