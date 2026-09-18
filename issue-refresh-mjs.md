## 背景

Excel-Ops 会持续增加 Agent 工作流、验证器、连接器和文档。用户通过 clone、fork 或本地工作目录使用项目时，目前没有统一入口判断本地版本是否落后，也无法在保留公司模板、Recipe、输入数据和输出文件的前提下安全更新系统能力。

项目需要一个面向 Agent 的 `refresh.mjs`：负责检查、预览和执行受控更新，但不参与 Excel 业务处理。Excel 读写仍由现有 Python 包完成。

## 目标

新增零依赖或最小依赖的更新入口：

```text
node refresh.mjs check
node refresh.mjs preview
node refresh.mjs apply --confirm
node refresh.mjs rollback
node refresh.mjs agent-check [--force]
```

它应当明确区分可更新的系统文件与用户文件，默认只检查，不得未经用户确认自动覆盖或安装。

## 建议行为

### `check`

- 比较本地版本、远端已测试版本及系统文件状态；
- 即使版本号相同，也能发现已测试快照中的系统文件变化；
- 返回机器可读结果，例如 `up-to-date`、`update-available`、`local-system-changed`、`offline`；
- 网络或 GitHub 不可用时 fail open，不阻止 Agent 继续处理本地工作簿。

### `preview`

- 列出准备新增、更新和删除的系统文件；
- 单独列出本地已经修改的系统文件与潜在冲突；
- 明确声明哪些用户文件不会被触碰；
- 不执行文件修改、依赖安装或 Git 操作。

### `apply --confirm`

- 只有显式确认后才应用更新；
- 更新前创建可恢复备份和状态记录；
- 使用允许列表或版本清单更新系统文件，不能依靠宽泛目录覆盖；
- 更新后运行最小 Doctor/测试检查；
- 检查失败时保留诊断信息并允许回滚。

### `rollback`

- 只回滚最近一次由 `refresh.mjs` 执行的系统更新；
- 不覆盖更新后新增或修改的用户文件；
- 回滚后重新运行完整性检查并返回机器可读结果。

### `agent-check`

- 每个本地 Agent 会话首次调用 Excel-Ops 时进行一次只读检查；
- 成功结果缓存 24 小时，`--force` 可以显式刷新；
- 仅提示有可用更新，不自动执行 `apply`；
- 不污染已有 JSON/CLI stdout 契约。

## 用户数据边界

在实现前先建立明确的 System/User Layer 清单。至少以下内容默认视为用户所有，不得自动更新或删除：

- 企业工作簿模板与 Template Profile；
- 项目 Recipe、人工确认与字段映射；
- 输入工作簿、复核包、Manifest、交付文件和输出目录；
- 云端连接器配置、凭据、客户数据与员工数据；
- 用户自行添加的本地示例或业务规则。

系统更新不得把这些内容上传到 GitHub、写入日志或复制进诊断包。

## 验收标准

- [ ] 提供 `check`、`preview`、显式确认的 `apply`、`rollback` 和只读 `agent-check`；
- [ ] 相同版本号下的系统文件变化仍可被发现；
- [ ] System/User Layer 使用显式清单或 Manifest，不通过宽泛目录猜测；
- [ ] `preview` 能展示文件级变化与本地冲突，但不修改文件；
- [ ] 未提供 `--confirm` 时不能应用更新；
- [ ] 更新前创建可验证备份，失败后能够安全回滚；
- [ ] Recipe、企业模板、数据、凭据和交付文件在 apply/rollback 前后保持不变；
- [ ] offline、无更新、有更新、本地系统文件已修改、缓存损坏等状态可区分；
- [ ] 缓存检查失败不会阻止正常 Excel-Ops 工作流；
- [ ] 日志和机器可读输出不包含 token、凭据或原始业务数据；
- [ ] Windows、macOS 和 Linux 路径行为有测试覆盖；
- [ ] 更新 README、Agent Skill 和中英文系统更新文档。

## 非目标

- 后台守护进程或操作系统定时任务；
- 未经用户同意自动更新；
- 自动安装 Python、Node 或系统级依赖；
- 自动解决用户修改与上游系统文件之间的冲突；
- 更新或迁移用户的企业模板和业务数据。

## 实现顺序建议

首个 PR 先完成 System/User Layer 清单、`check`、`preview` 和测试；第二个 PR 再实现备份、`apply`、`rollback` 与 Agent 首次调用提醒。这样可以先验证安全边界，再开放写操作。
