#!/usr/bin/env python3
"""Read input invoices and write an approval plan only under .runtime."""

from __future__ import annotations

import argparse
import secrets
from pathlib import Path
from typing import Any

from common import (
    DEFAULT_INPUT_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_PLAN_PATH,
    DEFAULT_PREVIEW_PATH,
    RUNTIME_DIR,
    assign_plan,
    input_snapshot,
    scan_input,
    utc_now,
    validate_path_layout,
    write_json,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="只读扫描 input，并在 .runtime 中生成预览和计划。")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_DIR)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_DIR)
    return parser


def _cell(value: Any) -> str:
    return str(value or "").replace("|", "\\|").replace("\n", " ")


def render_preview(plan: dict[str, Any]) -> str:
    lines = [
        "# 发票整理预览",
        "",
        f"- input：`{plan['input_dir']}`",
        f"- output：`{plan['output_dir']}`",
        f"- 扫描时间：`{plan['generated_at']}`",
        f"- 审批令牌：`{plan['approval_token']}`",
        "- 本次扫描只读取 input；未复制、改名、移动、删除或修改任何原始文件。",
        "",
        "## 拟输出文件",
        "",
        "| 序号 | 原文件名 | 拟生成文件名 | 费用类型 | 发票号码 | 票据类型 | 票面总金额 | 开票日期 | 状态/问题 |",
        "| ---: | --- | --- | --- | --- | --- | ---: | --- | --- |",
    ]
    for record in plan["records"]:
        if record.get("file_type") in {"archive", "archive_member"}:
            continue
        lines.append(
            "| {sequence} | {original} | {new} | {expense} | {number} | {document} | "
            "{total} | {date} | {status} |".format(
                sequence=_cell(record.get("sequence")),
                original=_cell(record.get("original_name")),
                new=_cell(record.get("new_name")),
                expense=_cell(record.get("expense_type")),
                number=_cell(record.get("invoice_number")),
                document=_cell(record.get("document_type")),
                total=_cell(record.get("total_amount")),
                date=_cell(record.get("invoice_date")),
                status=_cell(record.get("issue") or record.get("status")),
            )
        )

    lines += ["", "## ZIP 处理", ""]
    archives = [record for record in plan["records"] if record.get("file_type") == "archive"]
    if archives:
        for archive in archives:
            lines.append(
                f"- `{_cell(archive['original_name'])}`：只读提取计划中的 PDF 到 output；"
                "ZIP 保持原名、原位置和原内容，非 PDF 成员不输出。"
            )
    else:
        lines.append("- 无 ZIP")

    lines += ["", "## 待人工确认", ""]
    unresolved = [
        record for record in plan["records"] if record.get("status") == "needs_review"
    ]
    if unresolved:
        for record in unresolved:
            lines.append(
                f"- `{_cell(record.get('original_name'))}`："
                f"{_cell(record.get('issue') or '字段无法可靠识别')}"
            )
    else:
        lines.append("- 无")

    lines += [
        "",
        "## 确认后执行",
        "",
        "确认预览后运行：",
        "",
        f'`python scripts/organize_invoices.py --plan "{plan["plan_path"]}" '
        f'--confirm {plan["approval_token"]}`',
        "",
    ]
    return "\n".join(lines)


def main() -> int:
    args = build_parser().parse_args()
    try:
        input_dir, output_dir = validate_path_layout(args.input, args.output)
    except ValueError as exc:
        raise SystemExit(str(exc)) from exc
    if not input_dir.is_dir():
        raise SystemExit(f"input 目录不存在: {input_dir}")

    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    snapshot = input_snapshot(input_dir)
    records, ignored = scan_input(input_dir)
    records, issues = assign_plan(records, output_dir)
    plan = {
        "version": 2,
        "input_dir": str(input_dir),
        "output_dir": str(output_dir),
        "runtime_dir": str(RUNTIME_DIR.resolve()),
        "generated_at": utc_now(),
        "approval_token": secrets.token_urlsafe(18),
        "approved": False,
        "plan_path": str(DEFAULT_PLAN_PATH.resolve()),
        "preview_path": str(DEFAULT_PREVIEW_PATH.resolve()),
        "input_snapshot": snapshot,
        "ignored_files": ignored,
        "issues": issues,
        "records": records,
    }
    write_json(DEFAULT_PLAN_PATH, plan)
    DEFAULT_PREVIEW_PATH.write_text(render_preview(plan), encoding="utf-8")

    planned = sum(1 for record in records if record.get("new_name"))
    unresolved = sum(1 for record in records if record.get("status") == "needs_review")
    print(f"只读扫描完成: {len(snapshot)} 个 input 文件")
    print(f"拟输出票据: {planned}")
    print(f"待人工确认: {unresolved}")
    print(f"计划: {DEFAULT_PLAN_PATH}")
    print(f"预览: {DEFAULT_PREVIEW_PATH}")
    print(f"审批令牌: {plan['approval_token']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
