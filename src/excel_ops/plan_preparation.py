"""Prepare validated delivery-plan files from structured Agent intent."""

from __future__ import annotations

import hashlib
import json
import math
import os
import secrets
import stat
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator, Mapping

from .delivery import load_delivery_targets


class PlanPreparationError(ValueError):
    """Raised when a plan request would cross a safety boundary."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code


def prepare_delivery_plan(request: Mapping[str, Any]) -> dict[str, Any]:
    """Validate structured intent and atomically persist one delivery plan."""

    _validate_request_shape(request)
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


def _validate_request_shape(request: Mapping[str, Any]) -> None:
    if not isinstance(request, Mapping):
        raise PlanPreparationError("invalid_request", "request must be an object")
    for field in ("directory", "planPath"):
        if not isinstance(request.get(field), str):
            raise PlanPreparationError("invalid_request", f"{field} must be a string")
    for field in ("stagingDir", "deliveryDir", "recipe", "expectedDigest"):
        if field in request and not isinstance(request[field], str):
            raise PlanPreparationError("invalid_request", f"{field} must be a string")
    if "replace" in request and not isinstance(request["replace"], bool):
        raise PlanPreparationError("invalid_request", "replace must be a boolean")
    if "confidenceThreshold" in request:
        confidence = request["confidenceThreshold"]
        if (
            isinstance(confidence, bool)
            or not isinstance(confidence, (int, float))
            or not math.isfinite(confidence)
            or confidence < 0
            or confidence > 1
        ):
            raise PlanPreparationError(
                "invalid_request", "confidenceThreshold must be a finite number from 0 to 1"
            )
    for field in ("alsoAllow", "inputs", "targets"):
        value = request.get(field, [])
        if not isinstance(value, list):
            raise PlanPreparationError("invalid_request", f"{field} must be an array")
    for index, value in enumerate(request.get("alsoAllow", [])):
        if not isinstance(value, str):
            raise PlanPreparationError(
                "invalid_request", f"alsoAllow[{index}] must be a string"
            )
    for index, value in enumerate(request.get("inputs", [])):
        if not isinstance(value, str):
            raise PlanPreparationError("invalid_request", f"inputs[{index}] must be a string")
    for index, target in enumerate(request.get("targets", [])):
        if not isinstance(target, Mapping):
            raise PlanPreparationError(
                "invalid_request", f"targets[{index}] must be an object"
            )
        if "fieldColumns" in target and not isinstance(target["fieldColumns"], Mapping):
            raise PlanPreparationError(
                "invalid_request", f"targets[{index}].fieldColumns must be an object"
            )
        for field in ("aliases", "requiredFields", "periodExpectations"):
            if field in target and not isinstance(target[field], list):
                raise PlanPreparationError(
                    "invalid_request", f"targets[{index}].{field} must be an array"
                )
        for field in ("key", "template", "sheet", "recordIdField", "templateType", "deliveryName"):
            if field in target and not isinstance(target[field], str):
                raise PlanPreparationError(
                    "invalid_request", f"targets[{index}].{field} must be a string"
                )
        for field in ("aliases", "requiredFields"):
            for item_index, item in enumerate(target.get(field, [])):
                if not isinstance(item, str):
                    raise PlanPreparationError(
                        "invalid_request",
                        f"targets[{index}].{field}[{item_index}] must be a string",
                    )
        for field in ("headerRow", "dataStartRow", "styleSourceRow", "maxRows"):
            if field in target and (
                isinstance(target[field], bool) or not isinstance(target[field], int)
            ):
                raise PlanPreparationError(
                    "invalid_request", f"targets[{index}].{field} must be an integer"
                )
        if "formatPolicy" in target and not isinstance(target["formatPolicy"], Mapping):
            raise PlanPreparationError(
                "invalid_request", f"targets[{index}].formatPolicy must be an object"
            )
        for expectation_index, expectation in enumerate(target.get("periodExpectations", [])):
            if not isinstance(expectation, Mapping):
                raise PlanPreparationError(
                    "invalid_request",
                    f"targets[{index}].periodExpectations[{expectation_index}] must be an object",
                )


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


def _write_plan(
    path: Path, payload: bytes, *, replace: bool, expected_digest: Any
) -> None:
    with _verified_parent(path.parent) as directory_fd:
        with _plan_lock(path, directory_fd):
            _write_plan_locked(
                path,
                payload,
                replace=replace,
                expected_digest=expected_digest,
                directory_fd=directory_fd,
            )


def _write_plan_locked(
    path: Path,
    payload: bytes,
    *,
    replace: bool,
    expected_digest: Any,
    directory_fd: int | None,
) -> None:
    exists = _entry_exists(path, directory_fd)
    if not exists:
        try:
            descriptor = _open_entry(
                path,
                os.O_WRONLY | os.O_CREAT | os.O_EXCL,
                directory_fd,
                mode=0o600,
            )
        except FileExistsError as error:
            raise PlanPreparationError(
                "plan_already_exists", "plan appeared during creation"
            ) from error
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        return
    _require_regular_file(path, directory_fd)
    if not replace:
        raise PlanPreparationError("plan_already_exists", "refusing to overwrite an existing plan")
    if not str(expected_digest or "").strip():
        raise PlanPreparationError(
            "expected_digest_required", "replace requires the current plan digest"
        )
    before = _entry_digest(path, directory_fd)
    if before != expected_digest:
        raise PlanPreparationError("plan_changed", "existing plan digest does not match")

    temporary_name = f".{path.name}.{secrets.token_hex(12)}.tmp"
    temporary = path.parent / temporary_name
    descriptor = _open_entry(
        temporary,
        os.O_WRONLY | os.O_CREAT | os.O_EXCL,
        directory_fd,
        mode=0o600,
    )
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        if _entry_digest(path, directory_fd) != before:
            raise PlanPreparationError("plan_changed", "existing plan changed during preparation")
        _replace_entry(temporary, path, directory_fd)
    finally:
        _unlink_entry(temporary, directory_fd)


@contextmanager
def _verified_parent(parent: Path) -> Iterator[int | None]:
    """Hold a verified parent stable while the final entry is accessed."""

    if os.open not in os.supports_dir_fd or os.replace not in os.supports_dir_fd:
        with _hold_windows_path(parent):
            yield None
        return
    flags = os.O_RDONLY | getattr(os, "O_DIRECTORY", 0) | getattr(os, "O_NOFOLLOW", 0)
    descriptor = os.open(parent.anchor, flags)
    try:
        for component in parent.parts[1:]:
            next_descriptor = os.open(component, flags, dir_fd=descriptor)
            os.close(descriptor)
            descriptor = next_descriptor
        if not stat.S_ISDIR(os.fstat(descriptor).st_mode):
            raise PlanPreparationError("invalid_plan_directory", "planPath parent is not a directory")
        yield descriptor
    finally:
        os.close(descriptor)


@contextmanager
def _hold_windows_path(parent: Path) -> Iterator[None]:
    if os.name != "nt":
        raise PlanPreparationError(
            "secure_write_unavailable", "directory-relative secure writes are unavailable"
        )
    import ctypes
    from ctypes import wintypes

    create_file = ctypes.windll.kernel32.CreateFileW
    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE
    close_handle = ctypes.windll.kernel32.CloseHandle
    handles = []
    current = Path(parent.anchor)
    try:
        for part in parent.parts[1:]:
            current /= part
            handle = create_file(
                str(current),
                0x80,
                0x1 | 0x2,
                None,
                3,
                0x02000000 | 0x00200000,
                None,
            )
            if handle == wintypes.HANDLE(-1).value:
                raise PlanPreparationError(
                    "unsafe_plan_directory", f"could not lock verified directory {current.name}"
                )
            handles.append(handle)
        if parent.resolve(strict=True) != parent:
            raise PlanPreparationError("unsafe_plan_directory", "plan directory changed")
        yield
    finally:
        for handle in reversed(handles):
            close_handle(handle)


@contextmanager
def _plan_lock(path: Path, directory_fd: int | None) -> Iterator[None]:
    if os.name == "nt":
        with _windows_named_plan_lock(path):
            yield
        return
    if not _entry_exists(path, directory_fd):
        # Competing creators are serialized by O_CREAT | O_EXCL below.
        yield
        return
    import fcntl

    descriptor = _open_entry(path, os.O_RDONLY, directory_fd)
    try:
        if not stat.S_ISREG(os.fstat(descriptor).st_mode):
            raise PlanPreparationError("unsafe_plan_entry", "plan entry is not a regular file")
        fcntl.flock(descriptor, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(descriptor, fcntl.LOCK_UN)
    finally:
        os.close(descriptor)


@contextmanager
def _windows_named_plan_lock(path: Path) -> Iterator[None]:
    import ctypes
    from ctypes import wintypes

    name = "Local\\excel-ops-plan-" + hashlib.sha256(
        os.path.normcase(str(path)).encode("utf-8")
    ).hexdigest()
    create_mutex = ctypes.windll.kernel32.CreateMutexW
    create_mutex.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    create_mutex.restype = wintypes.HANDLE
    handle = create_mutex(None, False, name)
    if not handle:
        raise PlanPreparationError("plan_lock_failed", "could not create the plan lock")
    try:
        wait_result = ctypes.windll.kernel32.WaitForSingleObject(handle, 0xFFFFFFFF)
        if wait_result not in (0, 0x80):
            raise PlanPreparationError("plan_lock_failed", "could not acquire the plan lock")
        try:
            yield
        finally:
            ctypes.windll.kernel32.ReleaseMutex(handle)
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def _open_entry(path: Path, flags: int, directory_fd: int | None, *, mode: int = 0o777) -> int:
    no_follow = getattr(os, "O_NOFOLLOW", 0)
    if directory_fd is not None:
        return os.open(path.name, flags | no_follow, mode, dir_fd=directory_fd)
    if os.name == "nt":
        return _open_windows_entry(path, flags, mode)
    return os.open(path, flags, mode)


def _open_windows_entry(path: Path, flags: int, mode: int) -> int:
    import ctypes
    import msvcrt
    from ctypes import wintypes

    create_file = ctypes.windll.kernel32.CreateFileW
    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE
    access_mode = flags & (os.O_WRONLY | os.O_RDWR)
    access = 0x80000000
    if access_mode in (os.O_WRONLY, os.O_RDWR):
        access |= 0x40000000
    if flags & os.O_CREAT and flags & os.O_EXCL:
        disposition = 1
    elif flags & os.O_CREAT:
        disposition = 4
    else:
        disposition = 3
    handle = create_file(
        str(path),
        access,
        0x1 | 0x2,
        None,
        disposition,
        0x80 | 0x00200000,
        None,
    )
    if handle == wintypes.HANDLE(-1).value:
        error = ctypes.windll.kernel32.GetLastError()
        if error in (80, 183):
            raise FileExistsError(error, os.strerror(error), str(path))
        if error in (2, 3):
            raise FileNotFoundError(error, os.strerror(error), str(path))
        raise OSError(error, os.strerror(error), str(path))
    metadata = os.lstat(path)
    if not stat.S_ISREG(metadata.st_mode):
        ctypes.windll.kernel32.CloseHandle(handle)
        raise PlanPreparationError("unsafe_plan_entry", "plan entry must not be a reparse point")
    descriptor = msvcrt.open_osfhandle(
        handle,
        access_mode | getattr(os, "O_BINARY", 0),
    )
    return descriptor


def _entry_exists(path: Path, directory_fd: int | None) -> bool:
    try:
        if directory_fd is not None:
            os.stat(path.name, dir_fd=directory_fd, follow_symlinks=False)
        else:
            os.lstat(path)
        return True
    except FileNotFoundError:
        return False


def _require_regular_file(path: Path, directory_fd: int | None) -> None:
    metadata = (
        os.stat(path.name, dir_fd=directory_fd, follow_symlinks=False)
        if directory_fd is not None
        else os.lstat(path)
    )
    if not stat.S_ISREG(metadata.st_mode):
        raise PlanPreparationError("invalid_plan_path", "planPath must name a regular file")


def _entry_digest(path: Path, directory_fd: int | None) -> str:
    descriptor = _open_entry(path, os.O_RDONLY, directory_fd)
    digest = hashlib.sha256()
    with os.fdopen(descriptor, "rb") as stream:
        if not stat.S_ISREG(os.fstat(stream.fileno()).st_mode):
            raise PlanPreparationError("unsafe_plan_entry", "plan entry is not a regular file")
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _replace_entry(source: Path, destination: Path, directory_fd: int | None) -> None:
    if directory_fd is not None:
        os.replace(
            source.name,
            destination.name,
            src_dir_fd=directory_fd,
            dst_dir_fd=directory_fd,
        )
    else:
        os.replace(source, destination)


def _unlink_entry(path: Path, directory_fd: int | None) -> None:
    try:
        if directory_fd is not None:
            os.unlink(path.name, dir_fd=directory_fd)
        else:
            path.unlink()
    except FileNotFoundError:
        pass
