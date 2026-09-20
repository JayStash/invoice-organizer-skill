"""Shared extraction, planning, and safety helpers for invoice organization."""

from __future__ import annotations

import hashlib
import io
import json
import re
import zipfile
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


SKILL_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_DIR = SKILL_ROOT / "input"
DEFAULT_OUTPUT_DIR = SKILL_ROOT / "output"
RUNTIME_DIR = SKILL_ROOT / ".runtime"
DEFAULT_PLAN_PATH = RUNTIME_DIR / ".invoice-organization-plan.json"
DEFAULT_PREVIEW_PATH = RUNTIME_DIR / "发票整理预览.md"
WORKBOOK_NAME = "发票整理清单.xlsx"

SCAN_SUFFIXES = {".pdf", ".zip", ".jpg", ".jpeg", ".png", ".ofd"}
CATEGORY_ORDER = {"大型交通": 0, "酒店": 1, "市内交通": 2, "其他": 3}
EXPENSE_TYPES = {"飞机车船费", "住宿费", "打车费", "其他费用"}
MONEY_FIELDS = (
    "amount_excluding_tax",
    "tax_amount",
    "fuel_surcharge",
    "civil_aviation_development_fund",
    "total_amount",
)
EXCEL_HEADERS = (
    "文件名",
    "发票号码",
    "费用类型",
    "票据类型",
    "不含税金额",
    "税率",
    "税额",
    "燃油附加费",
    "民航发展基金",
    "票面总金额",
    "开票日期",
    "备注",
)

NUMBER = r"([+-]?\d[\d,]*(?:\.\d+)?)"


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve_path(path: Path) -> Path:
    return path.expanduser().resolve()


def paths_overlap(first: Path, second: Path) -> bool:
    first = resolve_path(first)
    second = resolve_path(second)
    return first == second or first.is_relative_to(second) or second.is_relative_to(first)


def validate_path_layout(input_dir: Path, output_dir: Path) -> tuple[Path, Path]:
    input_dir = resolve_path(input_dir)
    output_dir = resolve_path(output_dir)
    runtime_dir = resolve_path(RUNTIME_DIR)
    if paths_overlap(input_dir, output_dir):
        raise ValueError("input 与 output 不能相同、互相包含或嵌套。")
    if paths_overlap(input_dir, runtime_dir):
        raise ValueError("input 不能与 Skill 的 .runtime 目录重叠或嵌套。")
    if paths_overlap(output_dir, runtime_dir):
        raise ValueError("output 不能与 Skill 的 .runtime 目录重叠或嵌套。")
    return input_dir, output_dir


def ensure_direct_child(path: Path, directory: Path) -> None:
    path = resolve_path(path)
    directory = resolve_path(directory)
    if path.parent != directory:
        raise ValueError(f"源文件必须直接位于 input 中: {path}")


def parse_amount(value: str | None) -> str | None:
    if value is None:
        return None
    cleaned = value.replace(",", "").replace("￥", "").replace("¥", "").strip()
    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?", cleaned):
        return None
    try:
        number = Decimal(cleaned)
    except InvalidOperation:
        return None
    normalized = format(number, "f")
    if "." in normalized:
        normalized = normalized.rstrip("0").rstrip(".")
    return normalized


def decimal_or_none(value: Any) -> Decimal | None:
    normalized = parse_amount(str(value)) if value is not None else None
    return Decimal(normalized) if normalized is not None else None


def normalize_date(year: str, month: str, day: str) -> str:
    return date(int(year), int(month), int(day)).isoformat()


def _date_for_label(text: str, labels: tuple[str, ...]) -> str | None:
    joined = "|".join(re.escape(label) for label in labels)
    patterns = (
        rf"(?:{joined})\s*[:：]?\s*(20\d{{2}})[-/年]\s*(\d{{1,2}})[-/月]\s*(\d{{1,2}})日?",
        rf"(?:{joined})[^\d\n]{{0,20}}(20\d{{2}})(\d{{2}})(\d{{2}})",
    )
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            try:
                return normalize_date(*match.groups())
            except ValueError:
                continue
    return None


def invoice_date_from_text(text: str) -> str | None:
    return _date_for_label(text, ("开票日期",))


def all_dates_from_text(text: str) -> list[str]:
    values: list[str] = []
    for pattern in (
        r"(?<!\d)(20\d{2})年\s*(\d{1,2})月\s*(\d{1,2})日",
        r"(?<!\d)(20\d{2})[-/](\d{1,2})[-/](\d{1,2})(?!\d)",
    ):
        for match in re.finditer(pattern, text):
            try:
                values.append(normalize_date(*match.groups()))
            except ValueError:
                continue
    return sorted(set(values))


def business_date_from_text(text: str, category: str, invoice_date: str | None) -> str | None:
    if category == "大型交通":
        explicit = _date_for_label(
            text,
            ("出行日期", "乘车日期", "乘机日期", "开车日期", "航班日期"),
        )
        if explicit:
            return explicit
        candidates = all_dates_from_text(text)
        non_invoice = [value for value in candidates if value != invoice_date]
        return (non_invoice or candidates or [None])[0]
    if category == "市内交通":
        explicit = _date_for_label(text, ("行程起止日期", "出行日期", "乘车日期"))
        if explicit:
            return explicit
        candidates = all_dates_from_text(text)
        non_invoice = [value for value in candidates if value != invoice_date]
        return (non_invoice or candidates or [None])[0]
    if category == "酒店":
        explicit = _date_for_label(text, ("入住日期", "住宿日期", "消费日期"))
        return explicit or invoice_date
    return invoice_date


def extract_pdf_text(data: bytes) -> str:
    try:
        import pdfplumber  # type: ignore

        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception:
        try:
            from pypdf import PdfReader  # type: ignore

            reader = PdfReader(io.BytesIO(data))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception as exc:
            raise RuntimeError(f"PDF text extraction failed: {exc}") from exc


def invoice_number_from_text(text: str) -> str | None:
    match = re.search(r"发\s*票\s*号\s*码\s*[:：]?\s*(\d{8,})", text)
    return match.group(1) if match else None


def document_type_from_text(text: str, *, report: bool = False) -> str | None:
    if report and ("滴滴" in text or "DIDI" in text.upper()):
        return "滴滴出行-行程单"
    compact = re.sub(r"\s+", "", text)
    for title in (
        "电子发票（增值税专用发票）",
        "电子发票(增值税专用发票)",
        "电子发票（普通发票）",
        "电子发票(普通发票)",
        "铁路电子客票",
    ):
        if title in compact:
            return title.replace("(", "（").replace(")", "）")
    return None


def _labeled_amount(text: str, labels: tuple[str, ...], *, max_gap: int = 24) -> str | None:
    joined = "|".join(re.escape(label) for label in labels)
    match = re.search(
        rf"(?:{joined})[^\d+\-\n]{{0,{max_gap}}}[¥￥]?\s*{NUMBER}",
        text,
        flags=re.IGNORECASE,
    )
    return parse_amount(match.group(1)) if match else None


def standard_vat_totals(text: str) -> tuple[str | None, str | None]:
    pattern = rf"(?:^|\n)\s*合\s*计\s*[¥￥]\s*{NUMBER}\s*[¥￥]\s*{NUMBER}"
    match = re.search(pattern, text, flags=re.IGNORECASE)
    if not match:
        return None, None
    return parse_amount(match.group(1)), parse_amount(match.group(2))


def amount_excluding_tax_from_text(text: str) -> str | None:
    explicit = _labeled_amount(text, ("不含税金额", "合计金额（不含税）", "合计金额(不含税)"))
    if explicit is not None:
        return explicit
    amount, _ = standard_vat_totals(text)
    return amount


def tax_amount_from_text(text: str) -> str | None:
    explicit = _labeled_amount(text, ("合计税额", "税额合计"))
    if explicit is not None:
        return explicit
    _, tax = standard_vat_totals(text)
    return tax


def tax_rate_from_text(text: str) -> tuple[str | None, bool]:
    if not re.search(r"税\s*率", text):
        return None, False
    rates = {
        f"{parse_amount(match.group(1))}%"
        for match in re.finditer(r"(?<!\d)(\d+(?:\.\d+)?)\s*%", text)
        if parse_amount(match.group(1)) is not None
    }
    if len(rates) == 1:
        return next(iter(rates)), False
    return None, len(rates) > 1


def total_amount_from_text(text: str, purpose: str, document_type: str | None) -> str | None:
    patterns = (
        rf"价\s*税\s*合\s*计[^\n]{{0,80}}?（?\s*小\s*写\s*）?\s*[:：]?\s*[¥￥]\s*{NUMBER}",
        rf"(?:合计金额|总金额|实付金额|应付金额)\s*[:：]?\s*[¥￥]?\s*{NUMBER}",
        rf"合\s*计\s*[:：]?\s*[¥￥]?\s*{NUMBER}\s*元",
    )
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            value = parse_amount(match.group(1))
            if value is not None:
                return value
    if document_type == "铁路电子客票" or purpose == "高铁":
        for pattern in (
            rf"票\s*价\s*[:：]?\s*[¥￥]?\s*{NUMBER}",
            rf"[¥￥]\s*{NUMBER}\s*(?:元)?[^\n]{{0,12}}票\s*价",
        ):
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                value = parse_amount(match.group(1))
                if value is not None:
                    return value
    return None


def fuel_surcharge_from_text(text: str) -> str | None:
    return _labeled_amount(text, ("燃油附加费", "燃油费"))


def civil_aviation_fund_from_text(text: str) -> str | None:
    return _labeled_amount(text, ("民航发展基金", "民航发展基金费"))


def letter_from_name(name: str) -> str | None:
    match = re.search(r"(?:发票|行程报销单|行程单)\s*([A-Z])", name, flags=re.IGNORECASE)
    return match.group(1).upper() if match else None


def classify(name: str, text: str) -> tuple[str, str, str | None]:
    haystack = f"{name}\n{text}".lower()
    if any(token in haystack for token in ("滴滴", "didi", "行程报销单", "didi travel")):
        if "行程报销单" in name or "didi travel" in haystack or "行程单" in name:
            return "市内交通", "滴滴出行行程报销单", "打车费"
        return "市内交通", f"滴滴发票{letter_from_name(name) or ''}".rstrip(), "打车费"
    if any(token in haystack for token in ("高铁", "动车", "电子客票", "铁路", "车次", "12306")):
        return "大型交通", "高铁", "飞机车船费"
    if any(token in haystack for token in ("机票", "航班", "航空", "代订机票", "经济舱")):
        return "大型交通", "机票", "飞机车船费"
    if any(token in haystack for token in ("住宿费", "酒店", "民宿", "住宿服务")):
        return "酒店", "酒店", "住宿费"
    if any(token in haystack for token in ("出租车", "网约车", "机场大巴", "地铁", "高速", "共享单车")):
        return "市内交通", "市内交通", "打车费"
    if any(token in haystack for token in ("保险", "人身保险", "保单")):
        return "其他", "保险", "其他费用"
    if any(token in haystack for token in ("餐费", "餐饮", "食品")):
        return "其他", "餐费", "其他费用"
    return "其他", "其他", None


def is_didi_report(name: str, text: str) -> bool:
    lowered = f"{name}\n{text}".lower()
    return "行程报销单" in name or "didi travel - trip table" in lowered


def is_invoice_text(text: str, document_type: str | None) -> bool:
    return bool(invoice_number_from_text(text)) or document_type in {
        "电子发票（增值税专用发票）",
        "电子发票（普通发票）",
        "铁路电子客票",
    }


def read_pdf_record(path_name: str, data: bytes, *, archive_source: str | None = None) -> dict[str, Any]:
    text = extract_pdf_text(data)
    category, purpose, expense_type = classify(path_name, text)
    report = is_didi_report(path_name, text)
    document_type = document_type_from_text(text, report=report)
    invoice_date = invoice_date_from_text(text)
    tax_rate, multiple_rates = tax_rate_from_text(text)
    business_date = business_date_from_text(text, category, invoice_date)
    remarks: list[str] = []
    if multiple_rates:
        remarks.append("存在多个税率")
    if archive_source:
        remarks.append("ZIP 内提取")
    if report:
        remarks.append("配套行程单")
    if report:
        file_type = "didi_report"
    elif is_invoice_text(text, document_type):
        file_type = "archive_pdf" if archive_source else "invoice_pdf"
    else:
        file_type = "other"
    return {
        "original_name": path_name,
        "file_type": file_type,
        "category": category,
        "purpose": purpose,
        "business_date": business_date,
        "invoice_date": invoice_date,
        "invoice_number": invoice_number_from_text(text),
        "expense_type": expense_type,
        "document_type": document_type,
        "amount_excluding_tax": amount_excluding_tax_from_text(text),
        "tax_rate": tax_rate,
        "tax_amount": tax_amount_from_text(text),
        "fuel_surcharge": fuel_surcharge_from_text(text),
        "civil_aviation_development_fund": civil_aviation_fund_from_text(text),
        "total_amount": total_amount_from_text(text, purpose, document_type),
        "remark": "；".join(remarks) if remarks else None,
        "letter": letter_from_name(path_name),
        "archive_source": archive_source,
        "source_sha256": sha256_bytes(data),
    }


def input_snapshot(input_dir: Path) -> list[dict[str, Any]]:
    return [
        {"name": path.name, "size": path.stat().st_size, "sha256": sha256_file(path)}
        for path in sorted(input_dir.iterdir(), key=lambda item: item.name.casefold())
        if path.is_file()
    ]


def scan_input(input_dir: Path) -> tuple[list[dict[str, Any]], list[str]]:
    records: list[dict[str, Any]] = []
    ignored: list[str] = []
    for path in sorted(input_dir.iterdir(), key=lambda item: item.name.casefold()):
        if not path.is_file():
            ignored.append(path.name)
            continue
        suffix = path.suffix.lower()
        if suffix not in SCAN_SUFFIXES:
            ignored.append(path.name)
            continue
        if suffix == ".zip":
            archive_hash = sha256_file(path)
            archive_record: dict[str, Any] = {
                "original_name": path.name,
                "source_path": str(path),
                "source_sha256": archive_hash,
                "file_type": "archive",
                "archive_source": None,
                "status": "source_only",
                "issue": None,
            }
            records.append(archive_record)
            pdf_count = 0
            try:
                with zipfile.ZipFile(path) as archive:
                    for info in archive.infolist():
                        if info.is_dir():
                            continue
                        data = archive.read(info.filename)
                        if info.filename.lower().endswith(".pdf"):
                            pdf_count += 1
                            member = read_pdf_record(
                                Path(info.filename).name,
                                data,
                                archive_source=path.name,
                            )
                            member.update(
                                {
                                    "archive_member": info.filename,
                                    "archive_sha256": archive_hash,
                                    "member_sha256": sha256_bytes(data),
                                    "source_path": str(path),
                                }
                            )
                            records.append(member)
                        else:
                            records.append(
                                {
                                    "original_name": f"{path.name}::{info.filename}",
                                    "file_type": "archive_member",
                                    "archive_source": path.name,
                                    "archive_member": info.filename,
                                    "archive_sha256": archive_hash,
                                    "member_sha256": sha256_bytes(data),
                                    "source_path": str(path),
                                    "status": "ignored",
                                    "issue": "ZIP 中非 PDF 成员，不进入 output",
                                }
                            )
                if pdf_count == 0:
                    archive_record["status"] = "needs_review"
                    archive_record["issue"] = "ZIP 中没有 PDF 发票"
            except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
                archive_record["status"] = "needs_review"
                archive_record["issue"] = f"ZIP 无法读取: {exc}"
        elif suffix == ".pdf":
            try:
                record = read_pdf_record(path.name, path.read_bytes())
                record["source_sha256"] = sha256_file(path)
                record["source_path"] = str(path)
                records.append(record)
            except Exception as exc:
                records.append(
                    {
                        "original_name": path.name,
                        "source_path": str(path),
                        "source_sha256": sha256_file(path),
                        "file_type": "other",
                        "status": "needs_review",
                        "issue": f"PDF 无法识别: {exc}",
                    }
                )
        else:
            records.append(
                {
                    "original_name": path.name,
                    "source_path": str(path),
                    "source_sha256": sha256_file(path),
                    "file_type": "other",
                    "status": "needs_review",
                    "issue": "当前版本不能可靠提取此文件类型",
                }
            )
    return records, ignored


def _critical_issue(record: dict[str, Any]) -> str | None:
    if record.get("file_type") not in {"invoice_pdf", "archive_pdf"}:
        if record.get("file_type") == "other":
            return record.get("issue") or "未能确认该 PDF 为可整理票据"
        return record.get("issue")
    missing = [
        label
        for label, key in (
            ("业务日期", "business_date"),
            ("发票号码", "invoice_number"),
            ("票面总金额", "total_amount"),
        )
        if record.get(key) in (None, "")
    ]
    return "缺少关键字段: " + "、".join(missing) if missing else record.get("issue")


def output_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        [record for record in records if record.get("new_name") and record.get("status") != "needs_review"],
        key=lambda record: (
            record.get("sequence", 10**9),
            1 if record.get("file_type") == "didi_report" else 0,
            record.get("new_name", "").casefold(),
        ),
    )


def assign_plan(records: list[dict[str, Any]], output_dir: Path) -> tuple[list[dict[str, Any]], list[str]]:
    issues: list[str] = []
    for record in records:
        issue = _critical_issue(record)
        if issue and record.get("file_type") not in {"archive", "archive_member"}:
            record["status"] = "needs_review"
            record["issue"] = issue

    invoices = [
        record
        for record in records
        if record.get("file_type") in {"invoice_pdf", "archive_pdf"}
        and record.get("status") != "needs_review"
    ]
    by_number: dict[str, list[dict[str, Any]]] = {}
    for record in invoices:
        by_number.setdefault(record["invoice_number"], []).append(record)
    for number, matches in by_number.items():
        if len(matches) > 1:
            issues.append(f"发票号码重复: {number}")
            for record in matches:
                record["status"] = "needs_review"
                record["issue"] = f"发票号码重复: {number}"

    invoices = [record for record in invoices if record.get("status") != "needs_review"]
    invoices.sort(
        key=lambda record: (
            CATEGORY_ORDER.get(record.get("category", "其他"), 3),
            record.get("business_date") or "9999-99-99",
            record.get("original_name", "").casefold(),
        )
    )
    for sequence, record in enumerate(invoices, start=1):
        record["sequence"] = sequence
        extension = ".pdf" if record["file_type"] == "archive_pdf" else Path(record["original_name"]).suffix.lower()
        record["new_name"] = (
            f"{sequence}、{record['purpose']}-{record['invoice_number']}-"
            f"{record['total_amount']}元{extension}"
        )
        record["output_path"] = str(output_dir / record["new_name"])
        record["output_sha256"] = record.get("member_sha256") or record["source_sha256"]
        record["status"] = "planned"

    reports = [record for record in records if record.get("file_type") == "didi_report"]
    invoices_by_letter: dict[str, list[dict[str, Any]]] = {}
    for record in invoices:
        if record.get("purpose", "").startswith("滴滴发票") and record.get("letter"):
            invoices_by_letter.setdefault(record["letter"], []).append(record)
    for report in reports:
        letter = report.get("letter")
        matches = invoices_by_letter.get(letter or "", [])
        if len(matches) != 1:
            report["status"] = "needs_review"
            report["issue"] = f"滴滴行程单无法唯一匹配发票: {letter or '缺少 A/B/C 标记'}"
            continue
        invoice = matches[0]
        report["sequence"] = invoice["sequence"]
        report["new_name"] = f"{invoice['sequence']}、滴滴出行行程报销单{letter}.pdf"
        report["output_path"] = str(output_dir / report["new_name"])
        report["output_sha256"] = report["source_sha256"]
        report["didi_relation"] = invoice["new_name"]
        invoice["didi_relation"] = report["new_name"]
        report["status"] = "planned"

    destinations: dict[str, list[dict[str, Any]]] = {}
    for record in output_records(records):
        destinations.setdefault(record["new_name"], []).append(record)
    for name, matches in destinations.items():
        if len(matches) > 1:
            issues.append(f"目标名称重复: {name}")
            for record in matches:
                record["status"] = "needs_review"
                record["issue"] = f"目标名称重复: {name}"
    return records, issues


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, value: dict[str, Any]) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
