# Excel 交付工具

插件自带纯 JavaScript Excel 生成器，在 Cindy 沙箱内生成标准 Open XML 工作簿，并通过当前调用的工作目录写入能力保存二进制 `.xlsx`。不需要 Python、Node 进程、第三方账号或联网下载依赖。

## 翻译

调用 `ghost_call`，`ghost_id` 为 `localization-expert`，`tool` 为 `deliver_translation_excel`。

```json
{"rows":[{"source":"延迟开服补偿","translation":"Delayed Server Launch Compensation"}]}
```

每条严格提供 `source` 和 `translation` 两个字符串。工具固定表头“原文、译文”，不接受自定义列名或额外备注列。语言质量由当前 Agent 负责，导出工具不另行翻译传入内容。

## 校对

调用 `ghost_call`，`ghost_id` 为 `localization-expert`，`tool` 为 `deliver_proofread_excel`。

```json
{"rows":[{"source":"登录","before":"Log in","after":"Log in","reason":"无需修改"}]}
```

每条严格提供 `source`、`before`、`after`、`reason` 四个字符串，对应“原文、修改前译文、修改后译文、修改原因”。所有条目按输入顺序保留；修改后译文填写完整内容，不用删除线或仅列改动片段。原因非空；没有原文时 `source` 留空并在原因中说明范围限制。

## 文件回执与失败

- 单次 1 至 1000 行，总文本最多 100 万 UTF-16 单位，单元格最多 32767 个单位，最终文件最多 12 MiB。超过限制先按条目分批，不能静默丢弃。
- 工具在写入成功后才返回 `format: xlsx`、文件路径、文件链接、实际列名及行数。`relative_path` 相对于本次会话工作目录；最终回复使用真实回执提供可点击链接。
- 文件中所有单元格均为文本，包含样式、冻结首行和筛选。原文空格、换行、数字样式、模板变量和 XML 符号按原样保留。
- `EXCEL_INPUT_INVALID` 表示输入结构、字段类型、长度或字符不可保存；修正具体问题，不随意删除原文字符。`EXCEL_WRITE_FAILED` 表示尚未获得保存成功回执；说明原因并解决工作目录写入问题，不报告任务已经完成。
- 固定片段检查工具仅辅助语言复核，不能替代任何一种文件交付工具。
