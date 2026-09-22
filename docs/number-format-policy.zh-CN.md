中文 | [English](number-format-policy.md)

# 语义化数字格式策略

Excel-Ops 将单元格存储数值与 Excel 显示格式分开。格式策略可以声明工作簿默认值和字段级覆盖，但不会舍入或替换写入单元格的真实数值。

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

目前支持 CAD、USD、CNY 和 EUR。字段可以覆盖工作簿默认币种。单独的 `$` 或 `¥` 无法唯一确定币种，因此会被拒绝。声明 `currency_field` 后，如果同一字段出现混合币种，或记录币种与显式规则冲突，交付计划会被阻断并进入格式策略歧义项，不会静默套用一种格式。

`defaults` 提供工作簿级规则；`fields` 中的每个字段都可以覆盖它，因此显式字段规则优先于工作簿默认值。

`display_decimals`（或 `decimals`）只控制 Excel 显示格式。`storage_decimals`、`calculation_decimals`、`rounding`、币种和负数样式是彼此独立的策略元数据，会进入变更日志和交付 Manifest。写入器不会执行破坏性舍入；保存后会重新打开工作簿，核对真实数值和格式，再进入后续交付验证。

超过 Excel 15 位有效数字可靠范围的整数会使验证失败，不会在有损数值写入后仍被声称为“已保留”；类似标识符的长数字应继续以文本存储。

负数样式支持 `standard`、`red`、`parentheses` 和 `accounting`。策略未声明的字段继续保留企业模板原有样式。
