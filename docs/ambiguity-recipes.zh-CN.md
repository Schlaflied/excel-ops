中文 | [English](ambiguity-recipes.md)

# 批量歧义确认与 Recipe

Excel-Ops 将相同字段、相同规则的问题合并为一个确认项，避免要求用户逐行作答。每项必须展示代表性样本、候选解释、影响行数、推荐项、推荐理由和置信度；推荐只是建议，不是事实。

用户可以选择候选项、保持 `unknown`，并指定作用域：

- `this-run`：仅在当前运行有效，不写入项目文件；
- `project`：写入无凭据的 `excel-ops-recipe-v1` JSON Recipe，供后续运行复用。

如果已保存的选择不再属于当前候选集合，规则将进入 `conflict`，必须重新确认。`pending`、`unknown` 和 `conflict` 都是正式交付的阻断状态，不能被静默升级为已确认。

`review_rows_from_ambiguities()` 可将字段级问题转换成现有 #18 Review Pack 行；`manifest_decisions()` 则输出决定值、状态和来源，供 #20 Manifest 使用。当前实现只提供契约，不提前实现模板写回或交付验证。
