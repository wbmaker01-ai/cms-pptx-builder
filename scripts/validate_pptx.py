#!/usr/bin/env python3
"""PPTX structure validator for cms-pptx-builder compatibility mode.

This checks package/XML structure only. It never claims that slides were visually rendered.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


@dataclass
class Check:
    name: str
    status: str
    detail: str


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _read_xml(archive: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    try:
        return ET.fromstring(archive.read(name))
    except (KeyError, ET.ParseError, ValueError):
        return None


def _names(archive: zipfile.ZipFile, prefix: str, pattern: str) -> List[str]:
    regex = re.compile(pattern)
    return sorted(name for name in archive.namelist() if name.startswith(prefix) and regex.fullmatch(name.rsplit("/", 1)[-1]))


def _slide_names(archive: zipfile.ZipFile) -> List[str]:
    return _names(archive, "ppt/slides/", r"slide\d+\.xml")


def _note_names(archive: zipfile.ZipFile) -> List[str]:
    return _names(archive, "ppt/notesSlides/", r"notesSlide\d+\.xml")


def _ratio_check(archive: zipfile.ZipFile) -> Check:
    root = _read_xml(archive, "ppt/presentation.xml")
    if root is None:
        return Check("aspect_ratio", "FAIL", "ppt/presentation.xml을 읽을 수 없습니다.")
    size = next((node for node in root.iter() if _local(node.tag) == "sldSz"), None)
    if size is None:
        return Check("aspect_ratio", "FAIL", "슬라이드 크기(p:sldSz)를 확인할 수 없습니다.")
    try:
        cx, cy = float(size.attrib["cx"]), float(size.attrib["cy"])
        ratio = cx / cy
    except (KeyError, TypeError, ValueError, ZeroDivisionError):
        return Check("aspect_ratio", "FAIL", "슬라이드 크기 값이 올바르지 않습니다.")
    ok = abs(ratio - (16 / 9)) <= 0.02
    return Check("aspect_ratio", "PASS" if ok else "FAIL",
                 f"{cx:g} x {cy:g} EMU, 비율 {ratio:.4f} (16:9 {'일치' if ok else '불일치'})")


def _notes_checks(archive: zipfile.ZipFile, slide_count: int) -> List[Check]:
    notes = _note_names(archive)
    notes_match = len(notes) == slide_count and slide_count > 0
    result = [Check("notes.count", "PASS" if notes_match else "FAIL",
                    f"슬라이드 {slide_count}개, 발표자 노트 {len(notes)}개")]
    missing_sources: List[str] = []
    malformed: List[str] = []
    allowed = {"발표 스크립트", "교사 발문", "[Sources]", "Sources"}
    for name in notes:
        root = _read_xml(archive, name)
        text = "\n".join(node.text or "" for node in root.iter() if _local(node.tag) == "t") if root is not None else ""
        if "[Sources]" not in text:
            missing_sources.append(name)
        headings = re.findall(r"(?:^|\n)\s*\[([^\]]+)\]\s*", text)
        # Notes are expected to contain exactly three logical blocks. Accept plain
        # headings for the first two, while [Sources] must remain exact.
        if headings and any(h not in {"발표 스크립트", "교사 발문", "Sources"} for h in headings):
            malformed.append(name)
        if "발표 스크립트" not in text or "교사 발문" not in text:
            malformed.append(name)
    if not notes_match:
        result.append(Check("notes.sources", "FAIL", "발표자 노트 수가 슬라이드 수와 일치하지 않아 [Sources]를 검증할 수 없습니다."))
    elif missing_sources:
        result.append(Check("notes.sources", "FAIL", f"[Sources] 누락: {', '.join(missing_sources)}"))
    else:
        result.append(Check("notes.sources", "PASS", "모든 발표자 노트에 [Sources]가 있습니다."))
    if not notes_match:
        result.append(Check("notes.sections", "FAIL", "발표자 노트 수가 슬라이드 수와 일치하지 않아 3개 허용 항목을 검증할 수 없습니다."))
    elif malformed:
        result.append(Check("notes.sections", "WARN", f"3개 허용 항목 확인 필요: {', '.join(sorted(set(malformed)))}"))
    else:
        result.append(Check("notes.sections", "PASS", "발표 스크립트·교사 발문·[Sources] 항목을 확인했습니다."))
    return result


def _font_check(archive: zipfile.ZipFile) -> Check:
    fonts: set[str] = set()
    for name in archive.namelist():
        if not name.endswith(".xml"):
            continue
        root = _read_xml(archive, name)
        if root is None:
            continue
        for node in root.iter():
            for key, value in node.attrib.items():
                if _local(key) == "typeface" and value.strip():
                    fonts.add(value.strip())
    if not fonts:
        return Check("fonts", "WARN", "XML에서 사용 글꼴을 찾지 못했습니다.")
    pretendard = sorted(font for font in fonts if "Pretendard" in font)
    detail = ", ".join(sorted(fonts))
    return Check("fonts", "PASS" if pretendard else "WARN",
                 f"사용 글꼴: {detail}" + (" (Pretendard 확인)" if pretendard else " (Pretendard 확인 필요)"))


def _media_checks(archive: zipfile.ZipFile) -> List[Check]:
    chart_names = [name for name in archive.namelist() if re.fullmatch(r"ppt/charts/chart\d+\.xml", name)]
    workbook_names = [name for name in archive.namelist() if name.startswith("ppt/embeddings/")]
    charts_status = "PASS" if not chart_names or workbook_names else "WARN"
    charts_detail = f"차트 {len(chart_names)}개, 내장 통합문서 {len(workbook_names)}개"
    if chart_names and not workbook_names:
        charts_detail += "; 차트의 내장 통합문서를 확인하지 못했습니다."
    workbook_status = "PASS" if workbook_names or not chart_names else "WARN"
    workbook_detail = f"내장 통합문서 {len(workbook_names)}개"
    if chart_names and not workbook_names:
        workbook_detail += "; 차트가 있으면 데이터 내장 여부를 확인해야 합니다."
    return [Check("charts", charts_status, charts_detail),
            Check("embedded_workbooks", workbook_status, workbook_detail)]


def _external_check(archive: zipfile.ZipFile) -> Check:
    external: List[str] = []
    for name in archive.namelist():
        if not name.endswith(".rels"):
            continue
        root = _read_xml(archive, name)
        if root is None:
            continue
        for rel in root.iter():
            if _local(rel.tag) == "Relationship" and rel.attrib.get("TargetMode", "").lower() == "external":
                external.append(f"{name}: {rel.attrib.get('Target', '')}")
    return Check("external_relationships", "WARN" if external else "PASS",
                 "외부 관계 " + ("; ".join(external) if external else "없음"))


def validate(path: str) -> dict:
    checks: List[Check] = []
    if not os.path.isfile(path):
        return {"mode": "compatibility", "summary": "FAIL", "path": os.path.abspath(path),
                "checks": [asdict(Check("file", "FAIL", "파일이 존재하지 않습니다."))],
                "visual_rendering": "WARN: 시각 렌더링 여부는 확인하지 않음"}
    checks.append(Check("file", "PASS", f"파일 존재: {os.path.abspath(path)}"))
    try:
        archive = zipfile.ZipFile(path)
        bad_file = archive.testzip()
    except (OSError, zipfile.BadZipFile) as exc:
        checks.append(Check("zip", "FAIL", f"정상 PPTX ZIP이 아닙니다: {exc}"))
        return {"mode": "compatibility", "summary": "FAIL", "path": os.path.abspath(path),
                "checks": [asdict(c) for c in checks],
                "visual_rendering": "WARN: 시각 렌더링 여부는 확인하지 않음"}
    with archive:
        checks.append(Check("zip", "PASS" if bad_file is None else "FAIL",
                            "ZIP 항목을 읽을 수 있습니다." if bad_file is None else f"손상된 ZIP 항목: {bad_file}"))
        required = ["[Content_Types].xml", "ppt/presentation.xml"]
        missing = [name for name in required if name not in archive.namelist()]
        checks.append(Check("pptx_structure", "PASS" if not missing else "FAIL",
                            "필수 PPTX 구조를 확인했습니다." if not missing else f"필수 항목 누락: {', '.join(missing)}"))
        slides = _slide_names(archive)
        checks.append(Check("slides.count", "PASS" if slides else "FAIL", f"슬라이드 {len(slides)}개"))
        checks.append(_ratio_check(archive))
        checks.extend(_notes_checks(archive, len(slides)))
        checks.append(_font_check(archive))
        checks.extend(_media_checks(archive))
        checks.append(_external_check(archive))
    statuses = [check.status for check in checks]
    summary = "FAIL" if "FAIL" in statuses else ("WARN" if "WARN" in statuses else "PASS")
    return {"mode": "compatibility", "summary": summary, "path": os.path.abspath(path),
            "checks": [asdict(c) for c in checks],
            "visual_rendering": "WARN: 시각 렌더링 완료 여부는 이 검사에서 추정하지 않음"}


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="cms-pptx-builder PPTX 구조 검증")
    parser.add_argument("pptx", help="검사할 .pptx 경로")
    parser.add_argument("--json", action="store_true", help="JSON으로 출력")
    args = parser.parse_args(argv)
    result = validate(args.pptx)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"PPTX 호환 모드 검증: {result['summary']}")
        for check in result["checks"]:
            print(f"- {check['status']}: {check['name']} — {check['detail']}")
        print(f"- {result['visual_rendering']}")
    return 1 if result["summary"] == "FAIL" else 0


if __name__ == "__main__":
    raise SystemExit(main())
