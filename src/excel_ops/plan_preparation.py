"""Prepare validated delivery-plan files from structured Agent intent."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
from pathlib import Path
from typing import Any, Mapping

from .delivery import load_delivery_targets


class PlanPreparationError(ValueError):
    """Raised when a plan request would cross a safety boundary."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code


def prepare_delivery_plan(request: Mapping[str, Any]) -> dict[str, Any]:
    """Validate structured intent and atomically persist one delivery plan."""

    root = _existing_directory(request.get("directory"), "directory")
    allowed = (root,) + tuple(
        _existing_directory(item, "alsoAllow") for item in request.get("alsoAllow", ())
    )
    plan_path = _resolve_under(request.get("planPath"), root, allowed, "planPath")
    if not plan_path.parent.is_dir():
        raise PlanPreparationError(
            "missing_plan_directory", "planPath parent must already exist"
        )

    review = _review_items(request)
    if review:
        return {
            "prepared": False,
            "status": "needs_review",
            "review": review,
            "next_tool": "excel_ops.prepare_delivery",
        }

    inputs = [
        _resolve_existing_file(item, root, allowed, "input")
        for item in request.get("inputs", ())
    ]
    targets = []
    for raw in request.get("targets", ()):
        template = _resolve_existing_file(raw["template"], root, allowed, "template")
        target: dict[str, Any] = {
            "key": str(raw["key"]).strip(),
            "aliases": list(raw.get("aliases", ())),
            "template": _portable_path(template, plan_path.parent),
            "sheet": str(raw["sheet"]).strip(),
            "header_row": raw.get("headerRow", 1),
            "data_start_row": raw.get("dataStartRow", 2),
            "field_columns": dict(raw["fieldColumns"]),
            "required_fields": list(raw.get("requiredFields", ())),
            "record_id_field": raw.get("recordIdField", "record_id"),
        }
        for source, destination in (
            ("templateType", "template_type"),
            ("styleSourceRow", "style_source_row"),
            ("maxRows", "max_rows"),
            ("periodExpectations", "period_expectations"),
            ("deliveryName", "delivery_name"),
            ("formatPolicy", "format_policy"),
        ):
            if source in raw and raw[source] is not None:
                target[destination] = raw[source]
        targets.append(target)

    staging = _resolve_under(
        request.get("stagingDir", "staging"), root, allowed, "stagingDir"
    )
    delivery = _resolve_under(
        request.get("deliveryDir", "delivery"), root, allowed, "deliveryDir"
    )
    payload: dict[str, Any] = {
        "inputs": [_portable_path(item, plan_path.parent) for item in inputs],
        "staging_dir": _portable_path(staging, plan_path.parent),
        "delivery_dir": _portable_path(delivery, plan_path.parent),
        "confidence_threshold": request.get("confidenceThreshold", 0.85),
        "targets": targets,
    }
    if request.get("recipe"):
        recipe = _resolve_existing_file(request["recipe"], root, allowed, "recipe")
        payload["recipe"] = _portable_path(recipe, plan_path.parent)

    try:
        load_delivery_targets(payload, base_dir=plan_path.parent)
    except (TypeError, ValueError) as error:
        raise PlanPreparationError("invalid_delivery_plan", str(error)) from error

    encoded = (json.dumps(payload, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    digest = hashlib.sha256(encoded).hexdigest()
    _write_plan(
        plan_path,
        encoded,
        replace=bool(request.get("replace", False)),
        expected_digest=request.get("expectedDigest"),
    )
    return {
        "prepared": True,
        "status": "prepared",
        "plan_path": str(plan_path),
        "plan_digest": digest,
        "summary": {
            "inputs": len(inputs),
            "targets": len(targets),
            "target_keys": [item["key"] for item in targets],
        },
        "next_tool": "excel_ops.plan_delivery",
    }


def _review_items(request: Mapping[str, Any]) -> list[dict[str, str]]:
    review: list[dict[str, str]] = []
    if not request.get("inputs"):
        review.append({"field": "inputs", "reason": "at least one input must be selected"})
    targets = request.get("targets")
    if not isinstance(targets, list) or not targets:
        review.append({"field": "targets", "reason": "at least one target must be declared"})
        return review
    for index, target in enumerate(targets):
        prefix = f"targets[{index}]"
        if not isinstance(target, Mapping):
            review.append({"field": prefix, "reason": "target must be an object"})
            continue
        for key, label in (
            ("key", "destination key"),
            ("template", "template path"),
            ("sheet", "target sheet"),
        ):
            if not str(target.get(key) or "").strip():
                review.append({"field": f"{prefix}.{key}", "reason": f"{label} is required"})
        if not isinstance(target.get("fieldColumns"), Mapping) or not target.get("fieldColumns"):
            review.append(
                {"field": f"{prefix}.fieldColumns", "reason": "an explicit field mapping is required"}
            )
    return review


def _existing_directory(value: Any, field: str) -> Path:
    if not str(value or "").strip():
        raise PlanPreparationError("missing_authorized_directory", f"{field} is required")
    try:
        path = Path(str(value)).resolve(strict=True)
    except OSError as error:
        raise PlanPreparationError("unreadable_authorized_directory", f"{field}: {error}") from error
    if not path.is_dir():
        raise PlanPreparationError("invalid_authorized_directory", f"{field} must be a directory")
    return path


def _resolve_under(value: Any, base: Path, allowed: tuple[Path, ...], field: str) -> Path:
    if not str(value or "").strip():
        raise PlanPreparationError("missing_path", f"{field} is required")
    raw = Path(str(value))
    candidate = (raw if raw.is_absolute() else base / raw).resolve(strict=False)
    if not any(candidate == root or root in candidate.parents for root in allowed):
        raise PlanPreparationError("path_outside_authorized_roots", f"{field} is outside authorized roots")
    return candidate


def _resolve_existing_file(
    value: Any, base: Path, allowed: tuple[Path, ...], field: str
) -> Path:
    candidate = _resolve_under(value, base, allowed, field)
    try:
        resolved = candidate.resolve(strict=True)
    except OSError as error:
        raise PlanPreparationError("missing_input_file", f"{field}: {error}") from error
    if not resolved.is_file():
        raise PlanPreparationError("invalid_input_file", f"{field} must be a file")
    if not any(resolved == root or root in resolved.parents for root in allowed):
        raise PlanPreparationError("path_outside_authorized_roots", f"{field} escapes authorized roots")
    return resolved


def _portable_path(path: Path, base: Path) -> str:
    return Path(os.path.relpath(path, base)).as_posix()


def _file_digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_plan(
    path: Path, payload: bytes, *, replace: bool, expected_digest: Any
) -> None:
    if not path.exists():
        try:
            descriptor = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        except FileExistsError as error:
            raise PlanPreparationError("plan_already_exists", "plan appeared during creation") from error
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        return
    if not path.is_file():
        raise PlanPreparationError("invalid_plan_path", "planPath must name a regular file")
    if not replace:
        raise PlanPreparationError("plan_already_exists", "refusing to overwrite an existing plan")
    if not str(expected_digest or "").strip():
        raise PlanPreparationError(
            "expected_digest_required", "replace requires the current plan digest"
        )
    before = _file_digest(path)
    if before != expected_digest:
        raise PlanPreparationError("plan_changed", "existing plan digest does not match")
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(dir=path.parent, delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        if _file_digest(path) != before:
            raise PlanPreparationError("plan_changed", "existing plan changed during preparation")
        os.replace(temporary, path)
        temporary = None
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
