"""Shared extraction, planning, and safety helpers for invoice organization."""

from __future__ import annotations

import hashlib
import io
import json
import re
import unicodedata
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
EXPENSE_TYPES = {"飞机车船费", "住宿费", "市内交通费", "其他", "其他费用"}
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

RIDE_BRAND_MARKERS = (
    ("滴滴", "滴滴"),
    ("didi", "滴滴"),
    ("高德", "高德"),
    ("amap", "高德"),
    ("曹操", "曹操"),
    ("风韵", "风韵"),
    ("首汽", "首汽"),
    ("t3出行", "T3出行"),
    ("享道", "享道"),
    ("阳光出行", "阳光出行"),
    ("如祺", "如祺"),
    ("美团打车", "美团打车"),
    ("花小猪", "花小猪"),
    ("神州专车", "神州专车"),
)
RIDE_MARKERS = (
    "滴滴", "didi", "高德", "amap", "曹操", "风韵", "首汽", "t3出行",
    "享道", "阳光出行", "如祺", "美团打车", "花小猪", "神州专车", "出租车",
    "网约车", "打车",
)
RIDE_REPORT_MARKERS = ("行程单", "行程报销单", "行程明细", "itinerary", "trip table", "trip-table")


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


def normalize_extracted_text(text: str) -> str:
    """Normalize compatibility glyphs emitted by some embedded PDF fonts."""
    return unicodedata.normalize("NFKC", text)


def _date_for_label(text: str, labels: tuple[str, ...]) -> str | None:
    joined = "|".join(re.escape(label) for label in labels)
    patterns = (
        rf"(?:{joined})\s*[:：]?\s*(20\d{{2}})\s*[-/年]\s*(\d{{1,2}})\s*[-/月]\s*(\d{{1,2}})\s*日?",
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
    normalized = normalize_extracted_text(text)
    labeled = _date_for_label(normalized, ("开票日期", "填开日期"))
    if labeled:
        return labeled
    compact = re.sub(r"\s+", "", normalized)
    if "电子发票" not in compact or "统一发票监制" not in compact:
        return None
    match = re.search(
        r"国家税务总局[^\d]{0,40}(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日",
        normalized,
    )
    if not match:
        return None
    try:
        return normalize_date(*match.groups())
    except ValueError:
        return None


def all_dates_from_text(text: str) -> list[str]:
    values: list[str] = []
    for pattern in (
        r"(?<!\d)(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日",
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
            ("出行日期", "出发日期", "乘车日期", "乘机日期", "开车日期", "航班日期"),
        )
        if explicit:
            return explicit
        candidates = all_dates_from_text(text)
        non_invoice = [value for value in candidates if value != invoice_date]
        return (non_invoice or candidates or [None])[0]
    if category == "市内交通":
        explicit = _date_for_label(
            text, ("行程起止日期", "行程日期", "行程时间", "出行日期", "乘车日期")
        )
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
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            return normalize_extracted_text(text)
    except Exception:
        try:
            from pypdf import PdfReader  # type: ignore

            reader = PdfReader(io.BytesIO(data))
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
            return normalize_extracted_text(text)
        except Exception as exc:
            raise RuntimeError(f"PDF text extraction failed: {exc}") from exc


def invoice_number_from_text(text: str) -> str | None:
    normalized = normalize_extracted_text(text)
    match = re.search(r"发\s*票\s*号\s*码\s*[:：]?\s*(\d{8,})", normalized)
    if match:
        return match.group(1)
    compact = re.sub(r"\s+", "", normalized)
    if "电子发票" not in compact or "统一发票监制" not in compact:
        return None
    candidates = set(re.findall(r"(?<!\d)(\d{20})(?!\d)", normalized))
    return next(iter(candidates)) if len(candidates) == 1 else None


def ride_brand_from_text(name: str, text: str) -> str | None:
    haystack = f"{name}\n{text}".lower()
    for marker, brand in RIDE_BRAND_MARKERS:
        if marker in haystack:
            return brand
    return None


def is_ride_document(name: str, text: str) -> bool:
    haystack = f"{name}\n{text}".lower()
    return any(marker in haystack for marker in RIDE_MARKERS)


def is_ride_report(name: str, text: str) -> bool:
    haystack = f"{name}\n{text}".lower()
    if not any(marker in haystack for marker in RIDE_REPORT_MARKERS):
        return False
    return is_ride_document(name, text)


def document_type_from_text(text: str, *, report: bool = False) -> str | None:
    if report:
        return f"{ride_brand_from_text('', text) or '网约车'}-行程单"
    compact = re.sub(r"\s+", "", normalize_extracted_text(text))
    if "电子客票" in compact and any(
        token in compact for token in ("中国铁路", "12306", "车次")
    ):
        return "铁路电子客票"
    if "航空运输电子客票行程单" in compact or all(
        token in compact for token in ("承运人", "航班号", "民航发展基金", "电子客票号码")
    ):
        return "航空运输电子客票行程单"
    for title in (
        "电子发票（增值税专用发票）",
        "电子发票(增值税专用发票)",
        "电子发票（普通发票）",
        "电子发票(普通发票)",
        "铁路电子客票",
    ):
        if title in compact:
            return title.replace("(", "（").replace(")", "）")
    if "增值税电子普通发票" in compact:
        return "电子发票（普通发票）"
    if all(token in compact for token in ("发票代码", "发票号码", "校验码", "价税合计")) and (
        "电子支付标识" in compact
        or re.search(r"电.{0,8}子.{0,8}发.{0,8}票", compact[:160])
    ):
        return "电子发票（普通发票）"
    return None


def _aviation_ticket_amounts(text: str) -> dict[str, str]:
    compact = re.sub(r"[ \t]+", " ", text)
    header = re.search(
        r"票价\s+燃油附加费\s+增值税税率\s+增值税税额\s+民航发展基金\s+其他税费\s+合计",
        compact,
    )
    if not header:
        return {}
    for line in compact[header.end() :].splitlines():
        rate = re.search(r"(?<!\d)(\d+(?:\.\d+)?)\s*%", line)
        amounts = [
            parse_amount(match)
            for match in re.findall(r"(?:CNY|[¥￥])\s*([+-]?\d[\d,]*(?:\.\d+)?)", line, re.IGNORECASE)
        ]
        if rate and len(amounts) >= 6 and all(value is not None for value in amounts[:6]):
            return {
                "fare": amounts[0],
                "fuel_surcharge": amounts[1],
                "tax_rate": f"{parse_amount(rate.group(1))}%",
                "tax_amount": amounts[2],
                "civil_aviation_fund": amounts[3],
                "other_tax": amounts[4],
                "total_amount": amounts[5],
            }
    return {}


def _labeled_amount(text: str, labels: tuple[str, ...], *, max_gap: int = 24) -> str | None:
    joined = "|".join(re.escape(label) for label in labels)
    match = re.search(
        rf"(?:{joined})[^\d+\-\n]{{0,{max_gap}}}[¥￥]?\s*{NUMBER}",
        text,
        flags=re.IGNORECASE,
    )
    return parse_amount(match.group(1)) if match else None


def standard_vat_totals(text: str) -> tuple[str | None, str | None]:
    normalized = normalize_extracted_text(text)
    patterns = (
        rf"(?:^|\n)\s*合\s*计\s*[¥￥]\s*{NUMBER}\s*[¥￥]\s*{NUMBER}",
        rf"(?:^|\n)\s*[¥￥]\s*{NUMBER}\s*[¥￥]\s*{NUMBER}\s*\r?\n\s*合\s*计(?:\s|$)",
    )
    for pattern in patterns:
        match = re.search(pattern, normalized, flags=re.IGNORECASE)
        if match:
            return parse_amount(match.group(1)), parse_amount(match.group(2))
    return None, None


def amount_excluding_tax_from_text(text: str) -> str | None:
    aviation = _aviation_ticket_amounts(text)
    if aviation:
        return aviation["fare"]
    explicit = _labeled_amount(text, ("不含税金额", "合计金额（不含税）", "合计金额(不含税)"))
    if explicit is not None:
        return explicit
    amount, _ = standard_vat_totals(text)
    return amount


def tax_amount_from_text(text: str) -> str | None:
    aviation = _aviation_ticket_amounts(text)
    if aviation:
        return aviation["tax_amount"]
    explicit = _labeled_amount(text, ("合计税额", "税额合计"))
    if explicit is not None:
        return explicit
    _, tax = standard_vat_totals(text)
    return tax


def tax_rate_from_text(text: str) -> tuple[str | None, bool]:
    aviation = _aviation_ticket_amounts(text)
    if aviation:
        return aviation["tax_rate"], False
    if not re.search(r"税\s*率", text):
        return None, False
    rates = {
        f"{parse_amount(match.group(1))}%"
        for match in re.finditer(r"(?<!\d)(\d+(?:\.\d+)?)\s*%", text)
        if parse_amount(match.group(1)) is not None
    }
    if "不征税" in text:
        rates.add("不征税")
    if len(rates) == 1:
        return next(iter(rates)), False
    return None, len(rates) > 1


def total_amount_from_text(text: str, purpose: str, document_type: str | None) -> str | None:
    aviation = _aviation_ticket_amounts(text)
    if aviation:
        return aviation["total_amount"]
    patterns = (
        rf"价\s*税\s*合\s*计[^\n]{{0,100}}?[（(]\s*小\s*写\s*[）)][^\d+\-\n]{{0,12}}{NUMBER}",
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
            rf"(?:票[ \t]*价|退[ \t]*票[ \t]*费)[ \t]*[:：]?[ \t]*[¥￥]?[ \t]*{NUMBER}",
            rf"(?:票[ \t]*价|退[ \t]*票[ \t]*费)[ \t]*[:：]?[ \t]*\r?\n[ \t]*[¥￥][ \t]*{NUMBER}",
            rf"[¥￥][ \t]*{NUMBER}[ \t]*(?:元)?[ \t]*(?:\r?\n[ \t]*)?(?:票[ \t]*价|退[ \t]*票[ \t]*费)",
        ):
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                value = parse_amount(match.group(1))
                if value is not None:
                    return value
    return None


def fuel_surcharge_from_text(text: str) -> str | None:
    aviation = _aviation_ticket_amounts(text)
    if aviation:
        return aviation["fuel_surcharge"]
    return _labeled_amount(text, ("燃油附加费", "燃油费"))


def civil_aviation_fund_from_text(text: str) -> str | None:
    aviation = _aviation_ticket_amounts(text)
    if aviation:
        return aviation["civil_aviation_fund"]
    return _labeled_amount(text, ("民航发展基金", "民航发展基金费"))


def letter_from_name(name: str) -> str | None:
    match = re.search(r"(?:发票|行程报销单|行程单)\s*([A-Z])", name, flags=re.IGNORECASE)
    return match.group(1).upper() if match else None


def classify(name: str, text: str) -> tuple[str, str, str | None]:
    haystack = f"{name}\n{text}".lower()
    ride_brand = ride_brand_from_text(name, text)
    if is_ride_report(name, text):
        return "市内交通", f"{ride_brand or '网约车'}行程单", "市内交通费"
    if is_ride_document(name, text):
        return "市内交通", f"{ride_brand or '网约车'}发票", "市内交通费"
    if any(token in haystack for token in ("机票", "航班", "航空", "代订机票", "经济舱")):
        return "大型交通", "机票", "飞机车船费"
    if any(token in haystack for token in ("高铁", "动车", "电子客票", "铁路", "车次", "12306")):
        return "大型交通", "高铁", "飞机车船费"
    if any(token in haystack for token in ("船票", "客运港口", "水路客运", "轮渡", "客船")):
        return "大型交通", "船票", "飞机车船费"
    if any(token in haystack for token in ("住宿费", "酒店", "民宿", "住宿服务")):
        return "酒店", "酒店", "住宿费"
    if any(
        token in haystack
        for token in (
            "出租车", "网约车", "打车", "amap itinerary", "机场大巴", "地铁", "高速",
            "共享单车", "轨道交通", "公共交通", "乘车费",
        )
    ):
        return "市内交通", "市内交通", "市内交通费"
    if any(token in haystack for token in ("保险", "人身保险", "保单")):
        return "其他", "保险", "其他费用"
    if any(token in haystack for token in ("餐费", "餐饮", "食品")):
        return "其他", "餐费", "其他费用"
    if invoice_number_from_text(text) and any(
        token in haystack for token in ("项目名称", "货物或应税劳务", "服务名称")
    ):
        return "其他", "其他", "其他"
    return "其他", "其他", None


def is_didi_report(name: str, text: str) -> bool:
    """Backward-compatible alias for callers that used the old Didi helper."""
    return is_ride_report(name, text)


def is_invoice_text(text: str, document_type: str | None) -> bool:
    return bool(invoice_number_from_text(text)) or document_type in {
        "电子发票（增值税专用发票）",
        "电子发票（普通发票）",
        "铁路电子客票",
        "航空运输电子客票行程单",
    }


def read_pdf_record(path_name: str, data: bytes, *, archive_source: str | None = None) -> dict[str, Any]:
    text = extract_pdf_text(data)
    category, purpose, expense_type = classify(path_name, text)
    report = is_ride_report(path_name, text)
    ride_document = is_ride_document(path_name, text)
    ride_brand = ride_brand_from_text(path_name, text)
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
        file_type = "ride_report"
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
        "ride_brand": ride_brand,
        "ride_document": ride_document,
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
        [
            record
            for record in records
            if record.get("new_name")
            and (
                record.get("status") != "needs_review"
                or record.get("file_type") == "ride_report"
            )
        ],
        key=lambda record: (
            record.get("sequence", 10**9),
            1 if record.get("file_type") == "ride_report" else 0,
            record.get("new_name", "").casefold(),
        ),
    )


def _letter_for_index(index: int) -> str:
    letters = ""
    value = index
    while True:
        letters = chr(ord("A") + value % 26) + letters
        value = value // 26 - 1
        if value < 0:
            return letters


def _ride_report_match(
    report: dict[str, Any], invoices: list[dict[str, Any]]
) -> tuple[dict[str, Any] | None, str]:
    report_brand = report.get("ride_brand")
    candidates = [
        invoice
        for invoice in invoices
        if invoice.get("ride_document")
        and (not report_brand or invoice.get("ride_brand") == report_brand)
    ]
    if not candidates:
        return None, "未找到相同打车品牌的发票"
    if report_brand and len(candidates) == 1:
        return candidates[0], "品牌唯一匹配"
    scored: list[tuple[int, dict[str, Any]]] = []
    for invoice in candidates:
        score = 0
        if report.get("total_amount") and report.get("total_amount") == invoice.get("total_amount"):
            score += 4
        if report.get("business_date") and report.get("business_date") == invoice.get("business_date"):
            score += 3
        if report.get("invoice_date") and report.get("invoice_date") == invoice.get("invoice_date"):
            score += 1
        if report.get("letter") and invoice.get("ride_letter") == report.get("letter"):
            score += 5
        if score:
            scored.append((score, invoice))
    if not scored:
        return None, "缺少可证明唯一对应关系的金额、日期或字母信息"
    highest = max(score for score, _ in scored)
    matches = [invoice for score, invoice in scored if score == highest]
    if len(matches) != 1:
        return None, "金额、日期或字母信息无法唯一对应发票"
    return matches[0], "金额、日期或字母唯一匹配"


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

    ride_letter_counts: dict[str, int] = {}
    for record in invoices:
        if record.get("ride_document"):
            brand = record.get("ride_brand") or "网约车"
            index = ride_letter_counts.get(brand, 0)
            record["ride_letter"] = _letter_for_index(index)
            ride_letter_counts[brand] = index + 1
            record["purpose"] = f"{brand}发票{record['ride_letter']}"
            record["new_name"] = (
                f"{record['sequence']}、{record['purpose']}-{record['invoice_number']}-"
                f"{record['total_amount']}元{Path(record['new_name']).suffix}"
            )
            record["output_path"] = str(output_dir / record["new_name"])

    reports = [record for record in records if record.get("file_type") == "ride_report"]
    for report in reports:
        invoice, reason = _ride_report_match(report, invoices)
        if invoice is None:
            report["status"] = "needs_review"
            report["issue"] = f"打车行程单无法唯一匹配发票: {reason}"
            report["new_name"] = f"待人工复核、{Path(report['original_name']).name}"
            report["output_path"] = str(output_dir / report["new_name"])
            report["output_sha256"] = report.get("member_sha256") or report["source_sha256"]
            continue
        letter = invoice["ride_letter"]
        report["sequence"] = invoice["sequence"]
        report["ride_letter"] = letter
        report["ride_brand"] = invoice.get("ride_brand") or "网约车"
        report["purpose"] = f"{report['ride_brand']}行程单{letter}"
        report["new_name"] = f"{invoice['sequence']}、{report['purpose']}.pdf"
        report["output_path"] = str(output_dir / report["new_name"])
        report["output_sha256"] = report["source_sha256"]
        report["ride_relation"] = invoice["new_name"]
        invoice["ride_relation"] = report["new_name"]
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
