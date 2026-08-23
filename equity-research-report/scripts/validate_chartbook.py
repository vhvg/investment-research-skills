#!/usr/bin/env python3
"""Validate a lightweight equity-report chart workbook and its source URLs."""

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


def validate(path: Path, template: bool) -> list[dict[str, str]]:
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

            content_names = [name for name in names if name not in {"图表索引", "_图表模板"}]
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
                if not URL_RE.search(source_top):
                    issues.append(
                        issue(
                            "blocker",
                            "MISSING_SOURCE_URL",
                            f"工作表 {sheet_name} 前几行未填写原始链接。",
                        )
                    )

            if header and content_names:
                header_row, columns = header
                data_rows = {number: row for number, row in index_rows.items() if number > header_row}
                for sheet_name in content_names:
                    matching_rows = [
                        (number, row)
                        for number, row in data_rows.items()
                        if row.get(columns["工作表"], "") == sheet_name
                    ]
                    if not matching_rows:
                        issues.append(
                            issue("blocker", "UNMAPPED_SHEET", f"工作表 {sheet_name} 未在图表索引登记。")
                        )
                        continue
                    if len(matching_rows) > 1:
                        issues.append(
                            issue("blocker", "DUPLICATE_MAPPING", f"工作表 {sheet_name} 在图表索引重复登记。")
                        )
                    row_number, row = matching_rows[0]
                    missing = [
                        field
                        for field in ("编号", "报告章节", "图表标题", "状态")
                        if not row.get(columns[field], "")
                    ]
                    if missing:
                        issues.append(
                            issue(
                                "blocker",
                                "INCOMPLETE_MAPPING",
                                f"图表索引第 {row_number} 行缺少必要映射。",
                                "、".join(missing),
                            )
                        )
    except (zipfile.BadZipFile, ET.ParseError, KeyError) as exc:
        issues.append(issue("blocker", "INVALID_XLSX", "无法解析 XLSX 结构。", str(exc)))
    return issues


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    path = args.chartbook.expanduser().resolve()
    issues = validate(path, args.template)
    result = {
        "chartbook": str(path),
        "blockers": sum(item["level"] == "blocker" for item in issues),
        "warnings": sum(item["level"] == "warning" for item in issues),
        "issues": issues,
    }
    for item in issues:
        evidence = f" — {item['evidence']}" if item.get("evidence") else ""
        print(f"[{item['level'].upper()}] {item['code']}: {item['message']}{evidence}")
    if not issues:
        print("PASS: chart workbook mapping and top-of-sheet source links are complete")
    else:
        print(f"SUMMARY: {result['blockers']} blocker(s), {result['warnings']} warning(s)")
    if args.json_path:
        args.json_path.parent.mkdir(parents=True, exist_ok=True)
        args.json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return 1 if result["blockers"] or (args.strict and result["warnings"]) else 0


if __name__ == "__main__":
    raise SystemExit(main())
