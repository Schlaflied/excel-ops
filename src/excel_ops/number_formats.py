"""Semantic number formats that never rewrite the stored numeric value.

The policy is deliberately declarative.  It resolves a workbook default and
field overrides into Excel number-format codes, while keeping storage,
calculation, and display precision as separate pieces of provenance.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from numbers import Number
from typing import Any, Iterable, Mapping


SUPPORTED_KINDS = frozenset(
    {"amount", "exchange_rate", "hours", "headcount", "percentage", "tax_rate"}
)
SUPPORTED_CURRENCIES = frozenset({"CAD", "USD", "CNY", "EUR"})
AMBIGUOUS_CURRENCY_SYMBOLS = frozenset({"$", "¥", "￥"})

_CURRENCY_SYMBOLS = {"CAD": "CA$", "USD": "US$", "CNY": "¥", "EUR": "€"}
_DEFAULT_DECIMALS = {
    "amount": 2,
    "exchange_rate": 6,
    "hours": 2,
    "headcount": 0,
    "percentage": 2,
    "tax_rate": 2,
}


class NumberFormatPolicyError(ValueError):
    """Raised when a format policy cannot be applied without guessing."""

    def __init__(self, code: str, message: str, *, field_name: str | None = None):
        super().__init__(message)
        self.code = code
        self.field_name = field_name


@dataclass(frozen=True)
class FieldFormatPolicy:
    """A field-level semantic rule.

    ``storage_decimals`` and ``calculation_decimals`` are recorded constraints;
    they never cause destructive rounding in the writer. ``display_decimals``
    controls only the Excel number format.
    """

    kind: str
    display_decimals: int | None = None
    storage_decimals: int | None = None
    calculation_decimals: int | None = None
    currency: str | None = None
    currency_display: str = "symbol"
    negative_style: str = "standard"
    rounding: str = "none"

    def __post_init__(self) -> None:
        if self.kind not in SUPPORTED_KINDS:
            raise NumberFormatPolicyError("unsupported_format_kind", f"unsupported format kind: {self.kind}")
        for name in ("display_decimals", "storage_decimals", "calculation_decimals"):
            value = getattr(self, name)
            if value is not None and (not isinstance(value, int) or value < 0 or value > 15):
                raise NumberFormatPolicyError("invalid_precision", f"{name} must be an integer from 0 to 15")
        if self.currency_display not in {"symbol", "code"}:
            raise NumberFormatPolicyError("invalid_currency_display", "currency_display must be symbol or code")
        if self.negative_style not in {"standard", "red", "parentheses", "accounting"}:
            raise NumberFormatPolicyError(
                "invalid_negative_style",
                "negative_style must be standard, red, parentheses, or accounting",
            )
        if self.rounding not in {"none", "half_even", "half_up", "down"}:
            raise NumberFormatPolicyError("invalid_rounding", f"unsupported rounding policy: {self.rounding}")
        if self.currency is not None:
            _normalize_currency(self.currency)

    @property
    def decimals(self) -> int:
        return self.display_decimals if self.display_decimals is not None else _DEFAULT_DECIMALS[self.kind]

    @classmethod
    def from_mapping(cls, payload: Mapping[str, Any], *, field_name: str) -> "FieldFormatPolicy":
        kind = str(payload.get("kind") or _infer_kind(field_name))
        decimals = payload.get("display_decimals", payload.get("decimals"))
        style = str(payload.get("style") or "")
        negative_style = str(payload.get("negative_style") or ("accounting" if style == "accounting" else "standard"))
        return cls(
            kind=kind,
            display_decimals=_precision(decimals, "display_decimals"),
            storage_decimals=_precision(payload.get("storage_decimals"), "storage_decimals"),
            calculation_decimals=_precision(
                payload.get("calculation_decimals"), "calculation_decimals"
            ),
            currency=str(payload["currency"]) if payload.get("currency") is not None else None,
            currency_display=str(payload.get("currency_display", "symbol")),
            negative_style=negative_style,
            rounding=str(payload.get("rounding", "none")),
        )


@dataclass(frozen=True)
class FormatPolicy:
    """Workbook defaults plus explicit field overrides."""

    locale: str = "en-CA"
    currency: str | None = None
    currency_field: str | None = None
    fields: Mapping[str, FieldFormatPolicy] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.locale.strip():
            raise NumberFormatPolicyError("invalid_locale", "locale must not be empty")
        if self.currency is not None:
            _normalize_currency(self.currency)
        if not self.fields:
            raise NumberFormatPolicyError("missing_format_fields", "format policy needs at least one field")

    @classmethod
    def from_mapping(cls, payload: Mapping[str, Any]) -> "FormatPolicy":
        raw_fields = payload.get("fields")
        if not isinstance(raw_fields, Mapping) or not raw_fields:
            raise NumberFormatPolicyError("missing_format_fields", "format_policy.fields must be a non-empty object")
        raw_defaults = payload.get("defaults", {})
        if not isinstance(raw_defaults, Mapping):
            raise NumberFormatPolicyError("invalid_format_defaults", "format_policy.defaults must be an object")
        fields: dict[str, FieldFormatPolicy] = {}
        for name, raw in raw_fields.items():
            if not isinstance(raw, Mapping):
                raise NumberFormatPolicyError("invalid_format_field", f"format rule for {name} must be an object")
            fields[str(name)] = FieldFormatPolicy.from_mapping(
                {**raw_defaults, **raw}, field_name=str(name)
            )
        return cls(
            locale=str(payload.get("locale", "en-CA")),
            currency=str(payload["currency"]) if payload.get("currency") is not None else None,
            currency_field=str(payload["currency_field"]) if payload.get("currency_field") else None,
            fields=fields,
        )


@dataclass(frozen=True)
class ResolvedFieldFormat:
    field: str
    kind: str
    number_format: str
    display_decimals: int
    storage_decimals: int | None
    calculation_decimals: int | None
    currency: str | None
    currency_display: str
    negative_style: str
    rounding: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def resolve_format_policy(
    policy: FormatPolicy, records: Iterable[Mapping[str, Any] | object]
) -> dict[str, ResolvedFieldFormat]:
    """Resolve every configured field or fail closed on numeric/currency ambiguity."""

    rows = [_record_values(item) for item in records]
    resolved: dict[str, ResolvedFieldFormat] = {}
    for field_name, rule in policy.fields.items():
        values = [row[field_name] for row in rows if row.get(field_name) is not None]
        for value in values:
            if isinstance(value, bool) or not isinstance(value, Number):
                raise NumberFormatPolicyError(
                    "non_numeric_format_value",
                    f"field {field_name} contains a non-numeric value and cannot receive a numeric format",
                    field_name=field_name,
                )
        currency = _resolve_currency(policy, rule, rows, field_name)
        code = build_number_format(rule, currency=currency)
        resolved[field_name] = ResolvedFieldFormat(
            field=field_name,
            kind=rule.kind,
            number_format=code,
            display_decimals=rule.decimals,
            storage_decimals=rule.storage_decimals,
            calculation_decimals=rule.calculation_decimals,
            currency=currency,
            currency_display=rule.currency_display,
            negative_style=rule.negative_style,
            rounding=rule.rounding,
        )
    return resolved


def build_number_format(rule: FieldFormatPolicy, *, currency: str | None = None) -> str:
    """Build an Excel format code without touching a cell value."""

    zeros = "0" if rule.decimals == 0 else "0." + ("0" * rule.decimals)
    numeric = f"#,##{zeros}"
    if rule.kind in {"percentage", "tax_rate"}:
        positive = zeros + "%"
    elif rule.kind == "amount":
        if currency is None:
            raise NumberFormatPolicyError("currency_confirmation_required", "amount fields require an unambiguous currency")
        token = currency if rule.currency_display == "code" else _CURRENCY_SYMBOLS[currency]
        positive = f'"{token}"{numeric}'
    else:
        positive = numeric

    if rule.negative_style == "accounting":
        token = ""
        accounting_numeric = zeros + "%" if rule.kind in {"percentage", "tax_rate"} else numeric
        if rule.kind == "amount" and currency is not None:
            rendered = currency if rule.currency_display == "code" else _CURRENCY_SYMBOLS[currency]
            token = f'"{rendered}"'
        return (
            f'_({token}* {accounting_numeric}_);'
            f'_({token}* \\({accounting_numeric}\\);'
            f'_({token}* "-"??_);_(@_)'
        )
    if rule.negative_style == "standard":
        negative = f"-{positive}"
    elif rule.negative_style == "red":
        negative = f"[Red]-{positive}"
    else:
        negative = f"({positive})"
    return f"{positive};{negative};{positive}"


def policy_manifest(
    policy: FormatPolicy | None, resolved: Mapping[str, ResolvedFieldFormat] | None = None
) -> dict[str, Any]:
    """Return non-sensitive policy metadata suitable for a delivery Manifest."""

    if policy is None:
        return {}
    rules = resolved or {}
    return {
        "locale": policy.locale,
        "workbook_currency": _normalize_currency(policy.currency) if policy.currency else None,
        "currency_field": policy.currency_field,
        "fields": {
            name: rules[name].to_dict() if name in rules else {
                "field": name,
                "kind": rule.kind,
                "display_decimals": rule.decimals,
                "storage_decimals": rule.storage_decimals,
                "calculation_decimals": rule.calculation_decimals,
                "currency": _normalize_currency(rule.currency) if rule.currency else None,
                "currency_display": rule.currency_display,
                "negative_style": rule.negative_style,
                "rounding": rule.rounding,
            }
            for name, rule in sorted(policy.fields.items())
        },
    }


def _resolve_currency(
    policy: FormatPolicy,
    rule: FieldFormatPolicy,
    rows: list[Mapping[str, Any]],
    field_name: str,
) -> str | None:
    if rule.kind != "amount":
        return None
    declared = _normalize_currency(rule.currency) if rule.currency else None
    observed: set[str] = set()
    if policy.currency_field:
        for row in rows:
            if row.get(field_name) is None:
                continue
            if row.get(policy.currency_field) in (None, ""):
                if declared is None:
                    raise NumberFormatPolicyError(
                        "currency_confirmation_required",
                        f"field {field_name} has an amount with no currency",
                        field_name=field_name,
                    )
                continue
            observed.add(_normalize_currency(str(row[policy.currency_field])))
    if len(observed) > 1:
        raise NumberFormatPolicyError(
            "mixed_currency_confirmation_required",
            f"field {field_name} contains mixed currencies: {', '.join(sorted(observed))}",
            field_name=field_name,
        )
    observed_currency = next(iter(observed), None)
    if declared and observed_currency and declared != observed_currency:
        raise NumberFormatPolicyError(
            "currency_conflict_confirmation_required",
            f"field {field_name} declares {declared} but records contain {observed_currency}",
            field_name=field_name,
        )
    return declared or observed_currency or (_normalize_currency(policy.currency) if policy.currency else None)


def _normalize_currency(value: str) -> str:
    text = str(value).strip().upper()
    if text in AMBIGUOUS_CURRENCY_SYMBOLS:
        raise NumberFormatPolicyError(
            "ambiguous_currency_symbol",
            f"currency symbol {value!r} is ambiguous; use an ISO code such as CAD, USD, CNY, or EUR",
        )
    if text not in SUPPORTED_CURRENCIES:
        raise NumberFormatPolicyError("unsupported_currency", f"unsupported currency: {value}")
    return text


def _infer_kind(field_name: str) -> str:
    normalized = field_name.strip().lower()
    if any(token in normalized for token in ("exchange", "fx_rate", "currency_rate")):
        return "exchange_rate"
    if any(token in normalized for token in ("percent", "percentage", "pct")):
        return "percentage"
    if "tax" in normalized and "rate" in normalized:
        return "tax_rate"
    if any(token in normalized for token in ("hour", "hours")):
        return "hours"
    if any(token in normalized for token in ("headcount", "people", "persons")):
        return "headcount"
    if any(token in normalized for token in ("amount", "pay", "price", "cost", "total", "revenue")):
        return "amount"
    raise NumberFormatPolicyError(
        "format_kind_confirmation_required",
        f"cannot infer a semantic number-format kind for {field_name}; declare kind explicitly",
        field_name=field_name,
    )


def _precision(value: Any, name: str) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int):
        raise NumberFormatPolicyError(
            "invalid_precision", f"{name} must be an integer from 0 to 15"
        )
    return value


def _record_values(record: Mapping[str, Any] | object) -> Mapping[str, Any]:
    if isinstance(record, Mapping):
        return record
    try:
        return vars(record)
    except TypeError as exc:
        raise NumberFormatPolicyError("invalid_record", "records must be mappings or objects with fields") from exc
