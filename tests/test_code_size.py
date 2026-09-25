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
}


def _modules() -> dict[str, Path]:
    return {path.relative_to(SOURCE).as_posix(): path for path in sorted(SOURCE.rglob("*.py"))}


def _module_lines() -> dict[str, int]:
    return {
        name: len(path.read_text(encoding="utf-8").splitlines())
        for name, path in _modules().items()
    }


def _function_lines() -> dict[str, int]:
    sizes: dict[str, int] = {}

    def visit(node: ast.AST, module: str, scope: str) -> None:
        for child in ast.iter_child_nodes(node):
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                qualname = f"{scope}.{child.name}" if scope else child.name
                if not isinstance(child, ast.ClassDef):
                    key = f"{module}:{qualname}"
                    sizes[key] = max(sizes.get(key, 0), child.end_lineno - child.lineno + 1)
                visit(child, module, qualname)
            else:
                visit(child, module, scope)

    for name, path in _modules().items():
        visit(ast.parse(path.read_text(encoding="utf-8")), name, "")
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
