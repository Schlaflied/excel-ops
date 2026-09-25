"""Size ratchet: new code stays small, existing oversized code may only shrink.

Agents tend to keep appending to the file or function they are already in.
This test makes that visible: a new function over MAX_FUNCTION_LINES or a new
module over MAX_MODULE_LINES fails CI, and each entry below is a known
oversized item capped at its current length.  Split an item before adding to
it; once it is within the limit, delete its entry.
"""

from __future__ import annotations

import ast
from pathlib import Path

SOURCE = Path(__file__).resolve().parents[1] / "src" / "excel_ops"
MAX_FUNCTION_LINES = 100
MAX_MODULE_LINES = 800

OVERSIZED_FUNCTIONS = {
    "cli.py:main": 246,
    "delivery.py:run_delivery": 166,
    "delivery.py:_finish_run": 108,
    "delivery_write.py:_write_target": 137,
    "pdf_layout_verification.py:verify_pdf_layout": 139,
    "template_writer.py:write_template": 127,
}

OVERSIZED_MODULES = {
    "delivery.py": 834,
    "delivery_manifest.py": 807,
    "workdir.py": 940,
}


def _module_lines() -> dict[str, int]:
    return {path.name: len(path.read_text(encoding="utf-8").splitlines()) for path in SOURCE.glob("*.py")}


def _function_lines() -> dict[str, int]:
    sizes: dict[str, int] = {}
    for path in SOURCE.glob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                key = f"{path.name}:{node.name}"
                sizes[key] = max(sizes.get(key, 0), node.end_lineno - node.lineno + 1)
    return sizes


def _violations(sizes: dict[str, int], limit: int, allowed: dict[str, int], unit: str) -> list[str]:
    problems = []
    for key, size in sorted(sizes.items()):
        cap = allowed.get(key, limit)
        if size > cap:
            problems.append(
                f"{key} is {size} lines (cap {cap}); split it into named steps before adding to it"
            )
    for key, cap in sorted(allowed.items()):
        if sizes.get(key, 0) <= limit:
            problems.append(f"{key} is now within {limit} lines; delete its {unit} entry")
    return problems


def test_no_function_grows_past_its_cap() -> None:
    assert _violations(
        _function_lines(), MAX_FUNCTION_LINES, OVERSIZED_FUNCTIONS, "OVERSIZED_FUNCTIONS"
    ) == []


def test_no_module_grows_past_its_cap() -> None:
    assert _violations(
        _module_lines(), MAX_MODULE_LINES, OVERSIZED_MODULES, "OVERSIZED_MODULES"
    ) == []
