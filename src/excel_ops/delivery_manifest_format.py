"""Human-readable rendering of delivery Manifests beside the JSON contract."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Sequence

if TYPE_CHECKING:  # pragma: no cover - typing only
    from .delivery_manifest import DeliveryManifest


_STATUSES = ("written", "accepted", "review", "rejected", "skipped_existing")


def format_manifest(manifest: DeliveryManifest) -> str:
    """Render one Manifest for a human reviewer.

    Mirrors ``workdir.format_dry_run`` and ``idempotency.format_decision``: the
    JSON stays the machine contract, this is the same facts as readable lines.
    """

    period = (
        f"{manifest.period_start} .. {manifest.period_end}"
        if manifest.period_start or manifest.period_end
        else "not declared"
    )
    lines = [
        f"delivery manifest: {manifest.output_name}",
        f"  output           {manifest.output}",
        f"  output hash      {manifest.output_hash}",
        f"  destination      {manifest.destination_key}",
        f"  period           {period}",
        f"  template         {Path(manifest.template).name} ({manifest.template_version})",
        f"  recipe           {manifest.recipe_path or 'none'} ({manifest.recipe_version})",
        f"  verification     {manifest.verification.status}"
        f" ({manifest.verification.findings} finding(s))",
    ]
    if manifest.verification.codes:
        lines.append(f"  finding codes    {', '.join(manifest.verification.codes)}")
    if manifest.format_policy:
        formatted = manifest.format_policy.get("fields", {})
        lines.append(
            f"  number formats   locale={manifest.format_policy.get('locale')}, "
            f"fields={', '.join(sorted(formatted)) or 'none'}"
        )
    lines.append(
        "  run counts       "
        + ", ".join(f"{name}={manifest.run_counts.get(name, 0)}" for name in ("input", *_STATUSES))
    )
    lines.append(
        f"  this output      written={manifest.written}, rows in file={manifest.rows}, "
        f"review candidates={manifest.target_counts.get('review', 0)}, "
        f"skipped as existing={manifest.target_counts.get('skipped_existing', 0)}"
    )
    lines.append("  sources")
    for source in manifest.sources:
        lines.append(
            f"    - {source.file}: {source.records} record(s), "
            f"{source.written} written here, {source.hash}"
        )
    lines.append("  tabs")
    for tab in manifest.tabs:
        contributions = ", ".join(
            f"{Path(name).name}={count}" for name, count in tab.sources.items()
        )
        lines.append(
            f"    - {tab.sheet}: written={tab.written}, rows in file={tab.rows}"
            + (f", from {contributions}" if contributions else "")
        )
    if manifest.discrepancies:
        lines.append(f"  DISCREPANCIES    {', '.join(manifest.discrepancies)}")
    else:
        lines.append("  reconciled       yes (run counters match the persisted file)")
    if manifest.exports:
        lines.append("  exports")
        for item in manifest.exports:
            lines.append(
                f"    - {item.get('actual_format')}: {Path(str(item.get('output'))).name}, "
                f"sheets={','.join(item.get('sheets', ())) or 'none'}, "
                f"warnings={','.join(item.get('warnings', ())) or 'none'}, "
                f"digest={item.get('digest') or 'missing'}, "
                f"verification={item.get('verification') or 'unknown'}, "
                f"summary={item.get('summary') or 'none'}"
            )
    if manifest.formulas:
        lines.append(
            f"  formulas         status={manifest.formulas.get('status')}, "
            f"engine={manifest.formulas.get('engine') or 'not required'}, "
            f"rules={len(manifest.formulas.get('rules', ())) }"
        )
    lines.append(
        "  note             counts, hashes and metadata only; no cell value is recorded here"
    )
    return "\n".join(lines)


def format_manifests(manifests: Sequence[DeliveryManifest]) -> str:
    if not manifests:
        return "delivery manifest: nothing was delivered, so no manifest was generated"
    return "\n\n".join(format_manifest(item) for item in manifests)
