中文 | [English](refresh.md)

# 只读系统更新检查

Issue [#44](https://github.com/Schlaflied/excel-ops/issues/44) 分成两个 PR。首个 PR 只增加只读的系统文件清单与更新检查；工作簿处理仍由 Python 完成。

在仓库根目录运行：

```text
node refresh.mjs check
node refresh.mjs preview
```

两个命令各输出一个 JSON 对象，不修改文件。`check` 返回 `up-to-date`、`update-available`、`local-system-changed`、`offline`、`local-manifest-invalid` 或 `remote-manifest-invalid`。`preview` 还列出未来更新会新增、更新或删除的系统文件，以及本地冲突。网络不可用时返回 `offline`，不会阻断表格工作。

`refresh-manifest.json` 明确列出系统文件及其 SHA-256 哈希。远端比较以 GitHub `main` 上发布的清单为准；即使包版本号相同，文件变化仍能被发现。检查只读取清单所列的本地系统文件，不枚举用户文件；JSON 输出不包含文件内容或未知的用户路径。

企业模板、Template Profile、项目 Recipe、复核决定、输入工作簿、复核包、Manifest、输出文件、连接器配置、凭据，以及本地新增的示例或业务规则均属于用户层，不在系统清单中。若远端新增的系统路径已被本地文件占用，预览会把它列为冲突。

维护者新增系统文件时，应将其明确加入清单，然后运行 `node scripts/update-refresh-manifest.mjs` 更新哈希；验证时运行 `node scripts/update-refresh-manifest.mjs --check`。清单不计算自身的哈希。

`apply`、`rollback` 和 `agent-check` 留给 #44 的第二个 PR。这个 PR 不安装依赖、不调用 Git、不在检查时修改文件，也不改变 Python CLI 的 JSON 输出。
