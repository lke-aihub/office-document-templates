# Office Document Templates

用于生成格式严谨的中文公文和职场文案，并默认同时交付：

- 聊天中可直接复制的完整成稿；
- 与聊天成稿内容一致、按文种排版的 Word（`.docx`）文件。

Skill 调用名称：`$office-document-templates`

## 支持的文种

| 文种 | 脚本类型 | 默认排版档案 |
|---|---|---|
| 请假条 | `leave-note` | `simple-note` |
| 请示 | `request` | `business-formal` |
| 通知 | `notice` | `business-formal` |
| 申请书 | `application` | `business-formal` |
| 正式邮件 | `formal-email` | `email-archive` |
| 出差报告 | `travel-report` | `business-report` |
| 周报 | `weekly-report` | `business-report` |
| 月报 | `monthly-report` | `business-report` |
| 工作汇报 | `work-report` | `business-report` |
| 党政机关正式公文 | `official-document` | `official-gbt9704` |

普通企业请示、通知和申请不会仅因语气正式而套用党政机关公文版式。只有用户明确要求，或文件确属党政机关正式公文时，才使用 `official-gbt9704`。

## 使用方式

在 Codex 中直接说明需求即可，例如：

```text
使用 $office-document-templates 帮我写一份请假条：
明天去医院检查，请假两天，工作由小李暂时接手。
```

```text
使用 $office-document-templates 写一份南京客户拜访出差报告，
并同时生成 Word 文件。
```

```text
使用 $office-document-templates 起草一封正式的商务合作邀约邮件。
```

如果姓名、日期、金额、审批人、数据、成果或联系方式等必要信息缺失，Skill 会保留 `[待补充：字段]`，不会自行编造。

## 默认工作流程

1. 判断文种、对象、使用场景和语气。
2. 根据文种读取 `references/` 中对应的内容规范。
3. 使用用户提供的事实完成成稿，缺失信息保留待补充项。
4. 在聊天中输出完整成稿。
5. 将同一份内容写入 UTF-8 JSON，并生成 Word 文件。
6. 审计页面、核心字体和未处理的占位符。
7. 将 DOCX 渲染为页面图片并逐页检查，确认不存在截断、重叠、乱码或错误分页后交付。

## 目录结构

```text
office-document-templates/
├─ SKILL.md
├─ README.md
├─ agents/
│  └─ openai.yaml
├─ references/
│  ├─ official-forms.md
│  ├─ reports.md
│  ├─ emails.md
│  └─ format-standards.md
└─ scripts/
   ├─ generate_office_docx.py
   ├─ audit_office_docx.py
   └─ format_profiles.py
```

- `SKILL.md`：Skill 的入口说明、路由规则和交付要求。
- `agents/openai.yaml`：Codex UI 中显示的名称、简介和默认提示词。
- `references/`：不同文种的内容结构和排版依据。
- `scripts/generate_office_docx.py`：根据 JSON 内容生成确定性排版的 Word 文件。
- `scripts/audit_office_docx.py`：检查页面参数、核心字体和未解决的占位符。
- `scripts/format_profiles.py`：集中维护各类排版档案。

## 手动运行脚本

脚本应使用 Codex 工作区依赖加载器返回的 Python，不要直接依赖系统 Python。先查看目标文种的 JSON 结构：

```powershell
& '<workspace-python>' '<skill-dir>\scripts\generate_office_docx.py' `
  --print-schema travel-report
```

生成 Word：

```powershell
& '<workspace-python>' '<skill-dir>\scripts\generate_office_docx.py' `
  --type travel-report `
  --input '<payload.json>' `
  --output '<deliverable.docx>'
```

执行结构审计：

```powershell
& '<workspace-python>' '<skill-dir>\scripts\audit_office_docx.py' `
  '<deliverable.docx>' `
  --profile business-report
```

查看全部文种和排版档案：

```powershell
& '<workspace-python>' '<skill-dir>\scripts\generate_office_docx.py' --list-types
& '<workspace-python>' '<skill-dir>\scripts\generate_office_docx.py' --list-profiles
```

## 内容与交付约束

- 聊天成稿与 Word 正文中的事实、数字、附件名和结论必须一致。
- 不虚构姓名、日期、金额、审批人、数据、成果或承诺。
- 报告区分已完成、进行中、计划中和有风险事项。
- 正式邮件保留主题、收件人、抄送、正文、结束语、署名和附件清单。
- 用户提供旧稿时，保持事实和立场不变，只调整结构、语法、语气和排版。
- 用户明确要求不生成文件时，才省略 Word 交付。

## 维护说明

修改文种结构时，更新对应的 `references/*.md`；修改排版时，优先集中更新 `scripts/format_profiles.py`。脚本或入口说明发生变更后，应重新生成示例文档并完成结构审计与逐页视觉检查。
