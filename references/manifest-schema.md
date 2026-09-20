# 内部 Plan Schema

机器可读 JSON 是内部执行计划，不是最终用户清单。它只能位于 `.runtime/.invoice-organization-plan.json`，不能复制到 `output/`。最终用户清单只有 `output/发票整理清单.xlsx`。

## 顶层字段

| 字段 | 含义 |
| --- | --- |
| `version` | plan schema 版本 |
| `input_dir` | 只读输入目录绝对路径 |
| `output_dir` | 最终输出目录绝对路径 |
| `runtime_dir` | 当前 Skill 的内部运行目录 |
| `generated_at` | 扫描时间 |
| `approval_token` | 本次预览绑定的确认令牌 |
| `approved` | 是否成功执行 |
| `input_snapshot` | input 顶层全部文件的名称、大小和 SHA256 |
| `records` | 源文件、ZIP 成员、计划输出及提取字段 |
| `output_workbook_sha256` | 成功生成的 XLSX 哈希，用于幂等性和冲突检查 |
| `grand_total` | 非空 `total_amount` 的 Decimal 合计；全部为空时为 null |

## Record 字段

业务提取字段：

`invoice_number`、`expense_type`、`document_type`、`amount_excluding_tax`、`tax_rate`、`tax_amount`、`fuel_surcharge`、`civil_aviation_development_fund`、`total_amount`、`invoice_date`、`remark`。

内部字段可包括：

`original_name`、`new_name`、`source_path`、`output_path`、`file_type`、`category`、`purpose`、`business_date`、`archive_source`、`archive_member`、`didi_relation`、`source_sha256`、`archive_sha256`、`member_sha256`、`output_sha256`、`status`、`issue`。

`business_date` 仅用于排序；`invoice_date` 仅来自明确“开票日期”。两者不能互相替代。

## 状态

- `planned`：计划输出。
- `copied`：普通 PDF 已复制。
- `extracted`：ZIP PDF 已提取。
- `skipped_identical`：已存在且 SHA256 完全相同。
- `source_only`：ZIP 等只读来源，不进入 output。
- `ignored`：ZIP 非 PDF 成员或其他不输出内容。
- `needs_review`：不能可靠识别或配对，保留在 input 且不输出。

plan 可以包含绝对路径和哈希，因为它只用于内部审计。Excel 禁止包含这些字段。
