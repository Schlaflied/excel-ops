中文 | [English](refresh.md)

# 系统更新检查、更新与回滚

Issue [#44](https://github.com/Schlaflied/excel-ops/issues/44) 分成两个 PR。首个 PR 增加了只读的系统文件清单与 `check`、`preview`；第二个 PR 增加显式确认的 `apply`、针对该次更新的 `rollback`，以及带缓存的 `agent-check`。工作簿处理仍由 Python 完成。

在仓库根目录运行：

```text
node refresh.mjs check
node refresh.mjs preview
node refresh.mjs apply --confirm
node refresh.mjs rollback
node refresh.mjs agent-check [--force]
```

每个命令在 stdout 输出一个 JSON 对象。所有命令都不安装依赖、不调用 Git、不启动后台进程，也不改变 Python CLI 自身的输出。

## check 与 preview

两者都不修改文件。`check` 返回 `up-to-date`、`update-available`、`local-system-changed`、`offline`、`local-manifest-invalid` 或 `remote-manifest-invalid`。`preview` 还列出更新会新增、更新或删除的系统文件，以及本地冲突。网络不可用时返回 `offline`，不会阻断表格工作。

`refresh-manifest.json` 明确列出系统文件及其 SHA-256 哈希。远端比较以 GitHub `main` 上发布的清单为准；即使包版本号相同，文件变化仍能被发现。检查只读取清单所列的本地系统文件，不枚举用户文件；JSON 输出不包含文件内容或未知的用户路径。

## apply --confirm

未提供 `--confirm` 时返回 `confirmation-required`，退出码为 2，且不做任何改动：不下载、不写入、不备份、不删除。不存在部分应用或隐式应用。

提供 `--confirm` 后，更新按以下顺序执行：

1. 先执行 `preview` 比较。`offline`、`local-manifest-invalid`、`remote-manifest-invalid` 会在任何下载之前终止。
2. 若待更新路径上的本地系统文件与记录的哈希不一致，返回 `conflict-blocked`。冲突按文件列出，由你自行处理；`apply` 不会合并，也不会替你选一边。
3. 下载每个待更新的系统文件，并用发布清单校验 SHA-256。失败时在改动工作目录之前返回 `download-failed` 或 `download-verification-failed`。
4. 将所有即将被修改或删除的文件以及当前清单复制到 `.refresh/backups/<run-id>/files/`，随后写入状态记录 `.refresh/state.json`。状态记录在第一次写入之前就已存在，因此中断的更新仍可恢复。
5. 写入校验通过的文件，删除发布清单已移除的文件，并替换 `refresh-manifest.json`。
6. 运行最小 Doctor 检查：回读每个写入的文件并比对哈希，确认删除已生效，并重新校验本地清单。

成功时返回 `applied`，包含文件级的 `added`、`updated`、`removed` 与 `integrity.ok`。Doctor 检查或写入失败时返回 `apply-failed`，给出按文件的诊断信息，并把诊断保留在 `.refresh/state.json`，同时置 `rollbackAvailable`。只有发布清单中列出的路径会被写入，因此 `.refresh/` 和所有用户路径都不在更新范围内。

## rollback

`rollback` 只回滚 `apply` 自己记录的最近一次更新。它不是通用的恢复工具：只读取 `.refresh/state.json`，只处理其中列出的路径，并从该次运行的备份恢复其字节内容。

更新之后被改动过的文件保持原样，并在 `skipped` 中以 `changed-after-apply` 报告；更新之后新增的用户文件根本不在处理范围内。状态包括 `rolled-back`、`rolled-back-partial`（有跳过项）、`rolled-back-unverified`（完整性检查未通过）、`already-rolled-back`、`no-apply-recorded` 和 `state-invalid`。若某个路径无法写回（例如被占用锁定），会在 `skipped` 中以 `restore-failed` 报告，而不会中断整次回滚。部分成功的回滚在 `.refresh/state.json` 中记为 `rolled-back-partial`，因此在排除原因后可以再次运行 `rollback` 重试这些跳过项；已经恢复过的条目哈希与记录一致，重试只会写回相同的字节。只有完全成功的回滚才记为 `rolled-back`，再次运行时返回 `already-rolled-back`。恢复完成后会重新运行完整性检查，结果以 `integrity` 返回。

## agent-check

`agent-check` 是 Agent 首次调用 Excel-Ops 时的只读提醒，每个会话一次。它执行与 `check` 相同的比较，将成功结果缓存在 `.refresh/agent-check.json` 中 24 小时，并返回 `updateAvailable` 以及提示用户的建议。`--force` 显式跳过缓存。

它永远不会调用 `apply`。缓存缺失、过期、不可写或损坏时，会退回到一次新的只读检查；检查失败则返回 `offline`，因此不会阻断正常的 Excel-Ops 工作流。它的 JSON 是独立命令的输出，不会混入其他命令的 stdout 契约。

## 用户数据边界

企业模板、Template Profile、项目 Recipe、复核决定、确认后的字段映射、输入工作簿、复核包、Manifest、交付文件、输出目录、连接器配置、凭据、客户与员工数据，以及本地新增的示例或业务规则，均属于用户层。它们不在系统清单内，因此 `apply` 与 `rollback` 不会更新、删除、上传、写入日志，也不会复制进备份或诊断记录。备份只包含清单所列的系统文件。若清单新增的系统路径已被本地文件占用，会报告为冲突，而不是被覆盖。

维护者新增系统文件时，应将其明确加入清单，然后运行 `node scripts/update-refresh-manifest.mjs` 更新哈希；验证时运行 `node scripts/update-refresh-manifest.mjs --check`。清单不计算自身的哈希。`.refresh/` 只存放本地运行数据，已被 Git 忽略。
