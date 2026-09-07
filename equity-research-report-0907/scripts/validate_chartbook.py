#!/usr/bin/env python3
"""Statically check chartbook mappings, source locators, stored errors and report captions.

Does not recalculate formulas, verify source contents or reconcile arbitrary prose numbers.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_SHEET_RE = re.compile(r"^[FTI]\d{2}_.+")
CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")
URL_RE = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
LOCAL_SOURCE_RE = re.compile(
    r"本地材料[：:]\s*[^；;\n]+[；;]\s*定位[：:]\s*[^；;\n]+[；;]\s*证据[：:]\s*S\d+\b"
)
CAPTION_RE = re.compile(r"^(图表|图|表)\s*(\d+)\s*[：:]\s*(.+)$")

INDEX_HEADERS = (
    "编号",
    "报告章节",
    "工作表",
    "图表标题",
    "状态",
)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("chartbook", type=Path)
    parser.add_argument("--template", action="store_true", help="allow no populated F/T/I sheets")
    parser.add_argument("--strict", action="store_true", help="treat warnings as failure")
    parser.add_argument("--json", dest="json_path", type=Path)
    parser.add_argument("--report", type=Path, help="compare captions in a Markdown/text master draft")
    parser.add_argument("--report-fragment", action="store_true", help="allow indexed figures outside the supplied excerpt")
    return parser.parse_args(argv)


def issue(level: str, code: str, message: str, evidence: str | None = None) -> dict[str, str]:
    result = {"level": level, "code": code, "message": message}
    if evidence:
        result["evidence"] = evidence
    return result


def join_text(node: ET.Element) -> str:
    return "".join(text.text or "" for text in node.iter() if text.tag.endswith("}t"))


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [join_text(item) for item in root.findall(f"{{{NS_MAIN}}}si")]


def workbook_sheets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relation.attrib["Id"]: relation.attrib["Target"]
        for relation in relationships.findall(f"{{{NS_PACKAGE_REL}}}Relationship")
    }
    result: list[tuple[str, str]] = []
    for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
        rel_id = sheet.attrib[f"{{{NS_REL}}}id"]
        target = targets[rel_id].lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        result.append((sheet.attrib["name"], target))
    return result


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return join_text(cell)
    value = cell.find(f"{{{NS_MAIN}}}v")
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)]
        except (IndexError, ValueError):
            return value.text
    return value.text


def sheet_cells(
    archive: zipfile.ZipFile, path: str, shared_strings: list[str]
) -> tuple[dict[str, str], int]:
    root = ET.fromstring(archive.read(path))
    cells: dict[str, str] = {}
    formulas = len(root.findall(f".//{{{NS_MAIN}}}f"))
    for cell in root.findall(f".//{{{NS_MAIN}}}c"):
        reference = cell.attrib.get("r")
        if reference:
            cells[reference] = cell_value(cell, shared_strings).strip()
    return cells, formulas


def rows_from_cells(cells: dict[str, str]) -> dict[int, dict[str, str]]:
    rows: dict[int, dict[str, str]] = {}
    for reference, value in cells.items():
        match = CELL_RE.match(reference)
        if not match:
            continue
        column, row_text = match.groups()
        rows.setdefault(int(row_text), {})[column] = value
    return rows


def top_rows_text(cells: dict[str, str], max_row: int = 6) -> str:
    values: list[str] = []
    for reference, value in cells.items():
        match = CELL_RE.match(reference)
        if match and int(match.group(2)) <= max_row and value:
            values.append(value)
    return "\n".join(values)


def find_header(rows: dict[int, dict[str, str]]) -> tuple[int, dict[str, str]] | None:
    required = set(INDEX_HEADERS)
    for row_number, columns in rows.items():
        by_value = {value: column for column, value in columns.items() if value}
        if required.issubset(by_value):
            return row_number, {header: by_value[header] for header in INDEX_HEADERS}
    return None


def stored_cell_issues(archive: zipfile.ZipFile, path: str, name: str) -> list[dict[str, str]]:
    """Inspect saved XML only; caches can be stale even when present."""
    root = ET.fromstring(archive.read(path))
    issues = []
    uncached = []
    for cell in root.findall(f".//{{{NS_MAIN}}}c"):
        ref = f"{name}!{cell.attrib.get('r', '?')}"
        value = cell.find(f"{{{NS_MAIN}}}v")
        formula = cell.find(f"{{{NS_MAIN}}}f")
        if cell.attrib.get("t") == "e":
            issues.append(issue("blocker", "CELL_ERROR", "单元格保存了错误值。", f"{ref}: {value.text if value is not None else ''}"))
        if formula is not None:
            # Ignore error-looking literal strings used by IFERROR or display formulas.
            expression = re.sub(r'"(?:[^"]|"")*"', '', formula.text or '')
            if "#REF!" in expression.upper():
                issues.append(issue("blocker", "BROKEN_FORMULA_REFERENCE", "公式包含失效引用。", ref))
            if value is None or (value.text is None and cell.attrib.get("t") != "str"):
                uncached.append(ref)
    if uncached:
        issues.append(issue("warning", "FORMULA_CACHE_MISSING", f"{name} 有 {len(uncached)} 个公式缺少已保存结果，需重算并复核。", "、".join(uncached[:8])))
    return issues


def normalize_caption(value: str) -> str:
    return re.sub(r"\s+", "", value.replace("**", "").replace("__", "")).strip()


def compare_report(path: Path, records: list[dict[str, str]], fragment: bool) -> list[dict[str, str]]:
    if path.suffix.lower() not in {".md", ".markdown", ".txt"}:
        return [issue("blocker", "REPORT_FORMAT", "主稿对照仅支持 Markdown 或纯文本；DOCX/PDF 需核对其对应主稿。")]
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        return [issue("blocker", "REPORT_UNREADABLE", "无法读取主稿。", str(exc))]
    text = re.sub(r"(?s)<!--.*?-->", "", text)
    captions = []
    fence = None
    for line in text.splitlines():
        stripped = line.strip()
        marker = re.match(r"^(`{3,}|~{3,})", stripped)
        if marker:
            token = marker.group(1)
            if fence is None:
                fence = token
            elif token[0] == fence[0] and len(token) >= len(fence):
                fence = None
            continue
        if fence:
            continue
        stripped = re.sub(r"^#{1,6}\s+", "", stripped).replace("**", "").replace("__", "")
        match = CAPTION_RE.match(stripped)
        if match:
            kind, number, title = match.groups()
            captions.append((kind + str(int(number)), normalize_caption(title)))
    issues = []
    active = {normalize_caption(r["编号"]): r for r in records if r["状态"] != "停用"}
    inactive = {normalize_caption(r["编号"]) for r in records if r["状态"] == "停用"}
    seen = set()
    for number, title in captions:
        if number in seen:
            issues.append(issue("blocker", "DUPLICATE_REPORT_CAPTION", "主稿图表编号重复。", number))
        seen.add(number)
        if number not in active:
            code = "RETIRED_IN_REPORT" if number in inactive else "REPORT_FIGURE_UNMAPPED"
            issues.append(issue("blocker", code, "主稿图表没有有效索引映射。", number))
            continue
        record = active[number]
        if normalize_caption(record["图表标题"]) != title:
            issues.append(issue("blocker", "CAPTION_MISMATCH", "主稿图题与索引不一致。", number))
        if record["状态"] != "已完成":
            issues.append(issue("warning", "FIGURE_NOT_COMPLETE", "主稿使用的图表尚未标为已完成。", number))
    if not captions:
        issues.append(issue("warning", "NO_REPORT_CAPTIONS", "未识别到独立图题；请确认主稿范围及图题格式，不能据此认定无图表。"))
    if not fragment:
        for number in active.keys() - seen:
            issues.append(issue("blocker", "INDEX_FIGURE_NOT_IN_REPORT", "有效索引图表未出现在完整主稿中。", number))
    return issues


def validate(path: Path, template: bool, report: Path | None = None, report_fragment: bool = False) -> list[dict[str, str]]:
    issues: list[dict[str, str]] = []
    if path.suffix.lower() != ".xlsx":
        return [issue("blocker", "NOT_XLSX", "图表底稿必须为 .xlsx 文件。")]
    if not path.is_file():
        return [issue("blocker", "NOT_FOUND", "找不到图表底稿。", str(path))]

    try:
        with zipfile.ZipFile(path) as archive:
            bad_member = archive.testzip()
            if bad_member:
                return [issue("blocker", "CORRUPT_XLSX", "XLSX 压缩包校验失败。", bad_member)]

            shared_strings = read_shared_strings(archive)
            sheets = workbook_sheets(archive)
            if not sheets:
                return [issue("blocker", "NO_SHEETS", "工作簿没有工作表。")]

            names = [name for name, _ in sheets]
            if names[0] != "图表索引":
                issues.append(issue("blocker", "INDEX_NOT_FIRST", "第一张工作表必须为“图表索引”。", names[0]))
            if len(names) != len(set(names)):
                issues.append(issue("blocker", "DUPLICATE_SHEET", "存在重复工作表名称。"))

            payload: dict[str, tuple[dict[str, str], int]] = {}
            for name, sheet_path in sheets:
                try:
                    payload[name] = sheet_cells(archive, sheet_path, shared_strings)
                except KeyError:
                    issues.append(issue("blocker", "MISSING_SHEET_XML", "工作表 XML 缺失。", name))

            index_cells = payload.get("图表索引", ({}, 0))[0]
            index_rows = rows_from_cells(index_cells)
            header = find_header(index_rows)
            if not header:
                available = sorted({value for columns in index_rows.values() for value in columns.values() if value})
                issues.append(
                    issue(
                        "blocker",
                        "INDEX_HEADERS",
                        "图表索引缺少轻量模板的必备字段。",
                        "应包含：" + "、".join(INDEX_HEADERS) + "；当前字段：" + "、".join(available[:20]),
                    )
                )

            records = []
            if header:
                header_row, columns = header
                for number, row in index_rows.items():
                    if number <= header_row:
                        continue
                    record = {field: row.get(col, "") for field, col in columns.items()}
                    if not any(record.values()):
                        continue
                    records.append(record)
                    if record["状态"] == "停用":
                        continue
                    missing = [field for field, value in record.items() if not value]
                    if missing:
                        issues.append(issue("blocker", "INCOMPLETE_MAPPING", f"图表索引第 {number} 行缺少必要映射。", "、".join(missing)))
                    target = record["工作表"]
                    if target and (target not in names or target in {"图表索引", "_图表模板"}):
                        issues.append(issue("blocker", "INDEX_TARGET_MISSING", "有效索引行未指向存在的图表工作表。", target))
                for field, code in (("编号", "DUPLICATE_FIGURE_ID"), ("工作表", "DUPLICATE_MAPPING")):
                    seen = set()
                    for record in records:
                        if record["状态"] == "停用":
                            continue
                        value = normalize_caption(record[field]) if field == "编号" else record[field]
                        if value and value in seen:
                            issues.append(issue("blocker", code, f"有效索引的{field}重复。", value))
                        seen.add(value)
                active_ids = {normalize_caption(r["编号"]) for r in records if r["状态"] != "停用"}
                retired_ids = {normalize_caption(r["编号"]) for r in records if r["状态"] == "停用"}
                if active_ids & retired_ids:
                    issues.append(issue("blocker", "REUSED_RETIRED_ID", "有效图表使用了保留的停用编号。", "、".join(sorted(active_ids & retired_ids))))

            active_targets = {r["工作表"] for r in records if r["状态"] != "停用"}
            retired_targets = {r["工作表"] for r in records if r["状态"] == "停用"} - active_targets
            content_names = [name for name in names if name not in {"图表索引", "_图表模板"} and name not in retired_targets]
            for name, sheet_path in sheets:
                if name in payload and name not in retired_targets and name != "_图表模板":
                    issues.extend(stored_cell_issues(archive, sheet_path, name))
            if any(member.startswith("xl/externalLinks/externalLink") and member.endswith(".xml") for member in archive.namelist()):
                issues.append(issue("warning", "EXTERNAL_WORKBOOK_LINKS", "存在外部工作簿链接；需检查文件可用性、版本和重算结果。"))
            invalid_names = [name for name in content_names if not CONTENT_SHEET_RE.match(name)]
            if invalid_names:
                issues.append(
                    issue(
                        "blocker",
                        "SHEET_NAMING",
                        "图表工作表必须使用 Fnn_/Tnn_/Inn_ 前缀。",
                        "、".join(invalid_names),
                    )
                )
            if not content_names and not template:
                issues.append(issue("blocker", "NO_CONTENT_SHEETS", "成品底稿至少需要一个图表工作表。"))

            if "_图表模板" in names:
                template_top = top_rows_text(payload.get("_图表模板", ({}, 0))[0])
                if "数据来源" not in template_top:
                    issues.append(
                        issue(
                            "blocker",
                            "TEMPLATE_SOURCE_ROWS",
                            "_图表模板前几行必须保留数据来源填写位置。",
                        )
                    )

            for sheet_name in content_names:
                source_top = top_rows_text(payload.get(sheet_name, ({}, 0))[0])
                if "数据来源" not in source_top:
                    issues.append(
                        issue(
                            "blocker",
                            "SOURCE_LABEL_MISSING",
                            f"工作表 {sheet_name} 前几行未标明数据来源。",
                        )
                    )
                if LOCAL_SOURCE_RE.search(source_top):
                    issues.append(issue("warning", "LOCAL_SOURCE_REVIEW", f"工作表 {sheet_name} 使用本地材料；需实际核对文件、精确定位及证据台账。"))
                if not URL_RE.search(source_top) and not LOCAL_SOURCE_RE.search(source_top):
                    issues.append(
                        issue(
                            "blocker",
                            "MISSING_SOURCE_LOCATOR",
                            f"工作表 {sheet_name} 前六行缺少公开链接或完整本地材料定位。",
                        )
                    )

            if header and content_names:
                for sheet_name in content_names:
                    if sheet_name not in active_targets:
                        issues.append(
                            issue("blocker", "UNMAPPED_SHEET", f"工作表 {sheet_name} 未在图表索引登记。")
                        )
            if report is not None and header:
                issues.extend(compare_report(report, records, report_fragment))
    except (zipfile.BadZipFile, ET.ParseError, KeyError, OSError) as exc:
        issues.append(issue("blocker", "INVALID_XLSX", "无法解析 XLSX 结构。", str(exc)))
    return issues


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    path = args.chartbook.expanduser().resolve()
    if args.report_fragment and args.report is None:
        raise SystemExit("--report-fragment requires --report")
    issues = validate(path, args.template, args.report, args.report_fragment)
    result = {
        "chartbook": str(path),
        "blockers": sum(item["level"] == "blocker" for item in issues),
        "warnings": sum(item["level"] == "warning" for item in issues),
        "issues": issues,
        "limits": "Static checks only: no recalculation, cache freshness, source verification or arbitrary prose-number reconciliation.",
    }
    for item in issues:
        evidence = f" — {item['evidence']}" if item.get("evidence") else ""
        print(f"[{item['level'].upper()}] {item['code']}: {item['message']}{evidence}")
    if not issues:
        print("PASS: requested static chartbook checks passed")
    else:
        print(f"SUMMARY: {result['blockers']} blocker(s), {result['warnings']} warning(s)")
    print("LIMIT: no formula recalculation, source verification or arbitrary prose-number reconciliation")
    if args.json_path:
        args.json_path.parent.mkdir(parents=True, exist_ok=True)
        args.json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return 1 if result["blockers"] or (args.strict and result["warnings"]) else 0


if __name__ == "__main__":
    raise SystemExit(main())
