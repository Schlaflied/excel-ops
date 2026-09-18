"""Low-level cell addressing and write-safety predicates.

Column parsing and the "this cell is protected" predicate used to live in three
places (``template_writer``, ``delivery_verification`` and ``delivery``), each
with slightly different normalization.  They now share these helpers so a
planner's prediction and the writer's actual decision can never diverge.

Callers keep their own error types: these helpers raise plain ``ValueError`` and
each module converts it into whatever its contract promises.
"""

from __future__ import annotations

from typing import Any

from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string


def column_number(value: str | int) -> int:
    """Return the 1-based column index for a column letter or number.

    Letters are stripped and upper-cased before parsing so ``" b "`` and
    ``"B"`` mean the same column everywhere.  Raises ``ValueError`` for
    anything that is not a usable column reference.
    """

    if isinstance(value, bool):
        raise ValueError(f"invalid column: {value!r}")
    if isinstance(value, int):
        if value < 1:
            raise ValueError("column numbers start at 1")
        return value
    if not isinstance(value, str):
        raise ValueError(f"invalid column: {value!r}")
    try:
        return column_index_from_string(value.strip().upper())
    except (AttributeError, ValueError) as exc:
        raise ValueError(f"invalid column: {value}") from exc


def is_protected_formula_value(value: Any, *, overwrite_formulas: bool = False) -> bool:
    """Return whether a cell's current value must not be overwritten.

    This is the single definition of the writer's ``protected_formula`` skip:
    a cell already holding a formula is left alone unless the mapping opted in
    to overwriting formulas.  A planner can call it with a value read from the
    template ahead of time and reach the same verdict the writer will.
    """

    return isinstance(value, str) and value.startswith("=") and not overwrite_formulas


def is_merged_non_anchor(cell: Any) -> bool:
    """Return whether a cell is a merged-range follower rather than its anchor."""

    return isinstance(cell, MergedCell)
