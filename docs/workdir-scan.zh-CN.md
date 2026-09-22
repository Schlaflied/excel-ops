中文 | [English](workdir-scan.md)

# 工作目录扫描与本期文件选择

`scan_workdir(...)` 在任何表格工作开始之前回答一个问题：**这个工作目录里哪些文件是本期输入，哪些必须先由人看一眼？** 它对每个获授权的文件给出分类，说明每个文件被纳入或排除的原因，并且不移动、不改名、不写入、不删除任何东西。

```text
获授权的根目录（显式允许清单）
  -> 遍历，逐条重新校验是否仍在根目录内   （不跟随符号链接）
  -> 云同步残留检查                       （锁文件 / 未完成下载 / 冲突副本）
  -> 扩展名检查                           （不可读扩展名 -> exclude）
  -> 稳定窗口                             （不稳定 -> review，且完全不打开）
  -> 命名规则 -> 工作簿元数据 -> 修改时间
  -> 内容 hash 去重                       （字节相同的内容只处理一次）
  -> 疑似版本分组                         （名字像同一份，但不选出“真正那一份”）
  -> 用户覆盖 / 已保存的 workdir Recipe
  -> include / review / exclude，每个文件都带原因
```

## 保证

- **允许清单之外的文件不会被访问。** 扫描接受一份显式的授权根目录列表。是否仍在根目录内会对每一条目重新校验，而不是从父目录推断；符号链接被记为 `symlink_not_followed`，不会被跟随——与 `refresh.mjs` 对系统文件采用的根目录收敛写法一致。扩大范围必须是显式动作，绝不从目录结构推断。
- **正在写入或同步的文件不会被当作最终文件读取。** 只有文件大小和修改时间已经在稳定窗口内保持不变（并且在提供上一轮样本时与其一致），文件才会被打开。不稳定的文件不做 hash，没有 `content_hash`，以 `not_stable_yet` 进入 `review`。
- **重复内容按内容 hash 判定，而不是按文件名。** 字节完全相同但名字不同的文件组成一个 `DuplicateGroup`，其中一个代表保持 `include`，其余以 `duplicate_content` 排除。因为字节完全相同，这只是决定用哪个**名字**去工作，而不是判断哪份内容正确。
- **无法判断的最新版本不会被自动选择。** 内容不同的 `final.xlsx`、`final (1).xlsx`、`final-final.xlsx` 组成 `VersionCandidateGroup`，其 `selected` 恒为 `None`，全部成员进入 `review`。
- **云同步残留只被标记，不被吞进流程。** 锁文件（`~$…`）、未完成下载（`.crdownload`、`.part`、`.tmp`）和冲突副本（`… (conflicted copy from …)`、`…（冲突副本）`）不会进入 `include`。
- **用户覆盖可以改标签，但不能突破安全底线。** 覆盖不能授权允许清单之外的路径，不能把不稳定文件提升为 `include`（`override_refused_unstable_file`），也不能提升同步残留（`override_refused_sync_artifact`）。
- **扫描是只读的。** 本模块唯一会写的文件是 workdir Recipe，且只在显式调用 `save_workdir_recipe(...)` 时才写。

## 分类与处置

每个文件恰好落入一个分类和一个处置。

| 分类 | 含义 |
|---|---|
| `input` | 本期输入文件。 |
| `template` | 企业模板，依据扩展名、命名或工作簿元数据判定。 |
| `prior_delivery` | 以往交付件或已归档的历史文件。 |
| `review_return` | 回收的复核包。 |
| `unknown` | 没有决定性信号，交由人判断。 |

| 处置 | 含义 |
|---|---|
| `include` | 可以安全交给流程后续环节。 |
| `review` | 使用前必须由人确认。 |
| `exclude` | 本次运行刻意不使用，并记录原因。 |

产生该分类的信号记录在 `signal` 中（`name_pattern`、`workbook_metadata`、`modification_time`、`extension`、`stability_window`、`override` 或 `none`），因此每个判断都可复查。

## 用法

```python
from datetime import date
from excel_ops import PeriodWindow, ScanScope, format_dry_run, scan_workdir

scope = ScanScope.of(["/work/september", "/work/templates"], recursive=True)
result = scan_workdir(
    scope,
    period_window=PeriodWindow.of(date(2026, 9, 1), date(2026, 9, 30)),
    stability_window_seconds=5.0,
)

print(format_dry_run(result))            # 可读的 include / review / exclude 报告
print(result.counts())                   # {'include': 3, 'review': 4, 'exclude': 6}
for item in result.review:
    print(item.relative_path, item.reason, item.notes)
```

`period_window` 也可以来自已解析的业务周期：`PeriodWindow.from_period(resolve_period("2026年9月"))`。如果没有声明周期，就无法把数据文件与历史文件分开，因此这些文件以 `no_period_window_declared` 报为 `unknown` / `review`，不会被猜成 `input`。

要跨两轮判断文件是否仍在变化，把上一轮的观测结果传回去：

```python
first = scan_workdir(scope, period_window=window)
second = scan_workdir(scope, period_window=window, previous_samples=first.stability_samples())
```

### 命令行

```bash
excel-ops scan-workdir ./september --period-start 2026-09-01 --period-end 2026-09-30
excel-ops scan-workdir ./september --also-allow ./templates --no-recursive --json
excel-ops scan-workdir ./september --override "monthly.xlsx=template:exclude" --save-recipe workdir.json
excel-ops scan-workdir ./september --recipe workdir.json --result scan.json
```

这个命令本身就是 dry run：只报告会纳入什么、排除什么以及原因，不触碰任何文件。`--save-recipe` 是唯一会写文件的选项，而且只写 Recipe。

## 覆盖分类并保存为 Recipe

```python
from excel_ops import override_from_entry, save_workdir_recipe, scan_workdir

result = scan_workdir(scope, period_window=window)
entry = result.entry("northern-monthly.xlsx")
override = override_from_entry(entry, "template", disposition="exclude", note="regional template")
save_workdir_recipe([override], "recipes/workdir.json")

reused = scan_workdir(scope, period_window=window, recipe_path="recipes/workdir.json")
```

workdir Recipe 是独立的文件格式（`excel-ops-workdir-recipe-v1`），与字段级歧义决定使用的项目 Recipe（`excel-ops-recipe-v1`）分开。机制刻意保持同一形状——带版本的 JSON 文件、拒绝其他格式的 format 标记、成对的 `save_*` / `load_*`——但载荷以相对路径为键而不是 `field:question`，因为 workdir 的决定针对的是文件而不是数据字段。覆盖路径也可以是 `archive/*.xlsx` 这样的 `fnmatch` 通配符。相对路径始终使用 POSIX 分隔符，因此在 Windows 上写出的 Recipe 也能在别处复用。

## 尚未实现

- **未接入交付流程。** `run_delivery(...)` 和 `excel-ops deliver` 仍然使用显式声明的输入。把扫描结果的 `include` 集合喂给流程是后续独立的一步。
- **不移动、不改名、不删除任何文件。** 本模块只做分类和报告。
- **不解决云同步冲突。** 冲突副本会被识别并转入复核；判定哪一边胜出不在范围内，这与 `refresh.mjs` 的立场一致。
- **分类基于证据，不基于语义。** 模块读取文件名、扩展名、修改时间、内容 hash、工作表名和文档属性，不读取业务内容来判断文件“是什么意思”。
