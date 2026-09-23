[中文](number-format-policy.zh-CN.md) | English

# Semantic number-format policy

Excel-Ops separates the stored numeric value from its Excel display format. A
format policy can declare workbook defaults and field overrides without
rounding or replacing the value written to the cell.

```json
{
  "format_policy": {
    "locale": "en-CA",
    "currency": "CAD",
    "defaults": {"negative_style": "red", "rounding": "none"},
    "fields": {
      "gross_pay": {"kind": "amount", "style": "accounting", "decimals": 2},
      "exchange_rate": {"kind": "exchange_rate", "decimals": 6},
      "hours": {"kind": "hours", "decimals": 2},
      "headcount": {"kind": "headcount", "decimals": 0},
      "tax_rate": {"kind": "tax_rate", "decimals": 2}
    }
  }
}
```

Supported currencies are CAD, USD, CNY, and EUR. Currency fields can override
the workbook default. A bare `$` or `¥` is rejected because it does not identify
one currency. When `currency_field` is declared, mixed record currencies or a
conflict with an explicit currency block the delivery plan and appear as a
format-policy ambiguity instead of silently receiving one format.

`defaults` supplies workbook-level rule values. Each entry in `fields` can
override them; an explicit field rule therefore wins over a workbook default.

`display_decimals` (or `decimals`) controls only the Excel format code.
`storage_decimals`, `calculation_decimals`, `rounding`, currency, and negative
style are separate policy metadata recorded in the change log and delivery
Manifest. The writer never performs destructive rounding. It reopens the saved
workbook and verifies both the numeric values and applied formats before the
artifact can continue to delivery verification. Integers above Excel's
15-significant-digit reliability limit fail this verification instead of being
reported as preserved after a lossy numeric write; identifier-like values must
remain text.

Negative styles are `standard`, `red`, `parentheses`, and `accounting`. Fields
not named by the policy retain the enterprise template's existing style.
