---
name: office-document-templates
description: 生成请假条、出差报告、请示、通知、申请书、正式邮件、周报/月报、工作汇报与汇报材料等格式严谨的中文公文和职场文案，并默认同时在聊天中输出完整成稿、生成按文种排版的 Word（.docx）文件。用户提到这些文种、“公文”“汇报材料”“帮我写一份……”且任务明显需要固定格式（标题、称谓或主送机关、正文、落款、日期、主题或附件）时必须使用，即使未明说“按模板”或“生成Word”。也用于检查、改写、补全或统一此类文案格式；不用于纯文学创作、随意聊天或仅需一句非正式消息的请求。
---

# Office Document Templates

## 目标

一次完成两种交付：在聊天中给出可直接复制的完整成稿，同时生成内容一致、排版适合该文种的 `.docx`。优先保证事实准确、结构完整、语气得体和行动要求明确。

## 强制工作流

1. 判断文种、对象、场景、语气和是否属于党政机关正式公文。
2. 读取对应内容参考：
   - 请假条、请示、通知、申请书：读取 [references/official-forms.md](references/official-forms.md)。
   - 出差报告、周报/月报、工作汇报：读取 [references/reports.md](references/reports.md)。
   - 正式邮件：读取 [references/emails.md](references/emails.md)。
3. 需要核对字体或版式依据时读取 [references/format-standards.md](references/format-standards.md)；脚本已内置排版参数，正常生成时不要重新推导。
4. 提取用户事实并完成成稿。不得虚构姓名、日期、金额、审批人、数据、成果或承诺；少量缺失信息使用 `[待补充：字段]`。
5. 在聊天中输出完整成稿，不得只给文件链接或内容摘要。
6. 将同一成稿整理为 UTF-8 JSON，使用 [scripts/generate_office_docx.py](scripts/generate_office_docx.py) 生成 Word。默认不得手工重写 Word 排版。
7. 使用 [scripts/audit_office_docx.py](scripts/audit_office_docx.py) 检查页面和核心字体；再按文档能力执行 DOCX → PNG 渲染并逐页查看。发现截断、重叠、乱码、错误分页或格式漂移时，修正并重新渲染。
8. 最终同时交付聊天成稿和 `.docx`。仅当用户明确要求“不生成文件”时省略 Word。

## 选择文种和排版档案

| 文种 | `--type` | 默认档案 |
|---|---|---|
| 请假条 | `leave-note` | `simple-note` |
| 请示 | `request` | `business-formal` |
| 通知 | `notice` | `business-formal` |
| 申请书 | `application` | `business-formal` |
| 正式邮件 | `formal-email` | `email-archive` |
| 出差报告 | `travel-report` | `business-report` |
| 周报/月报 | `weekly-report` / `monthly-report` | `business-report` |
| 工作汇报 | `work-report` | `business-report` |
| 党政机关正式公文 | `official-document`，或其他类型加 `"official": true` | `official-gbt9704` |

只有确认文件属于党政机关正式公文或用户明确要求 `GB/T 9704-2012` 时使用 `official-gbt9704`。企业内部请示、通知和申请不得仅因语气正式而误用机关公文版式。

## 快速生成 Word

先使用 Codex 工作区依赖加载器取得 Python 路径，不使用系统 Python。用脚本打印一次目标文种示例结构，再据此写入任务临时 JSON：

```powershell
& '<workspace-python>' '<skill-dir>\scripts\generate_office_docx.py' --print-schema leave-note
& '<workspace-python>' '<skill-dir>\scripts\generate_office_docx.py' --type leave-note --input '<payload.json>' --output '<deliverable.docx>'
& '<workspace-python>' '<skill-dir>\scripts\audit_office_docx.py' '<deliverable.docx>' --profile simple-note
```

可用命令：

```powershell
& '<workspace-python>' '<skill-dir>\scripts\generate_office_docx.py' --list-types
& '<workspace-python>' '<skill-dir>\scripts\generate_office_docx.py' --list-profiles
```

把用户可见 `.docx` 写入当前任务允许的输出目录；JSON 和渲染图片放在任务临时目录，不作为交付物。文件名使用文种和事项，例如 `请假条-张三.docx`、`上海客户调研出差报告.docx`。

## Word 生成约束

- 聊天成稿与 Word 正文保持事实、数字、附件名和结论一致；不要在两处分别改写。
- 党政机关公文采用 A4、标准版心、二号小标宋标题、三号仿宋正文和四号宋体页码；层级依次为“一、”“（一）”“1.”“（1）”。
- 普通企业文案使用适合阅读和归档的商务版式，不强行套用机关公文字号。
- 正式邮件的 Word 文件作为归档版，保留主题、收件人、抄送、正文、结束语、署名和附件清单。
- 报告优先使用标题层级和简洁定义列表；只有真实可比的行列数据才使用表格。
- 保留 `[待补充：……]` 时，在聊天成稿后列出不超过 5 项待确认信息，并提示用户 Word 中也含占位符。
- 用户提供旧稿时保持事实和立场不变，只修正结构、语法、语气与排版。

## 写作规则

- 标题准确概括事项；开头交代目的或结论；正文一段一意；结尾明确请求、决定、计划或下一步。
- 报告按“结果—依据—问题—行动”组织，区分已完成、进行中和计划中事项。
- 邮件主题同时包含事项与期望动作；正文首段说明来意，末段给出动作与期限。
- 对外、向上或涉及人事、费用、合规时使用审慎措辞，不代替审批或作未经授权的承诺。
- 日期默认写作 `YYYY年M月D日`，人名、时间、金额、地点和附件名称前后一致。
- 删除重复背景、口号式表述、情绪化表达和无依据夸大。

## 提交前检查

- 文种、标题、称谓、署名、日期和附件是否匹配。
- 目的、事实、依据、请求或下一步是否明确。
- 聊天成稿与 Word 内容是否一致。
- Word 是否使用了正确档案并通过结构审计与逐页视觉检查。
- 收件人是否能一次读懂“发生了什么、需要做什么、何时完成”。
