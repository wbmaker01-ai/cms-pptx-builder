#!/usr/bin/env python3
"""cms-pptx-builder compatibility-mode environment diagnostics."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import platform
import shutil
import subprocess
import sys
from dataclasses import asdict, dataclass
from typing import Iterable, List, Optional, Tuple


@dataclass
class Check:
    name: str
    status: str
    detail: str
    path: Optional[str] = None


def _version(command: str) -> Optional[str]:
    try:
        result = subprocess.run([command, "--version"], capture_output=True,
                                text=True, timeout=4, check=False)
    except (OSError, subprocess.SubprocessError):
        return None
    text = (result.stdout or result.stderr or "").strip()
    return text.splitlines()[0][:240] if text else "available"


def _which(commands: Iterable[str]) -> Tuple[Optional[str], Optional[str]]:
    for command in commands:
        path = shutil.which(command)
        if path:
            return command, path
    return None, None


def _installed_powerpoint() -> Optional[str]:
    """Find PowerPoint installed on Windows even when POWERPNT is not on PATH."""
    if platform.system() != "Windows":
        return None
    candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft Office", "root", "Office16", "POWERPNT.EXE"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Microsoft Office", "root", "Office16", "POWERPNT.EXE"),
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft Office", "Office16", "POWERPNT.EXE"),
    ]
    try:
        import winreg
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for subkey in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE",
                           r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE"):
                try:
                    with winreg.OpenKey(hive, subkey) as key:
                        value, _ = winreg.QueryValueEx(key, None)
                        if value and os.path.isfile(value):
                            candidates.insert(0, value)
                except (FileNotFoundError, OSError):
                    continue
    except ImportError:
        pass
    return next((path for path in candidates if path and os.path.isfile(path)), None)


def _font_paths() -> List[str]:
    system = platform.system()
    if system == "Windows":
        roots = [os.path.join(os.environ.get("WINDIR", r"C:\\Windows"), "Fonts"),
                 os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")]
        names = ("Pretendard-Regular.ttf", "PretendardVariable.ttf", "malgun.ttf", "NanumGothic.ttf")
    elif system == "Darwin":
        roots = ["/Library/Fonts", os.path.expanduser("~/Library/Fonts")]
        names = ("Pretendard-Regular.ttf", "PretendardVariable.ttf", "AppleSDGothicNeo.ttc")
    else:
        roots = ["/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.local/share/fonts")]
        names = ("Pretendard-Regular.ttf", "PretendardVariable.ttf", "NotoSansCJK-Regular.ttc",
                 "NotoSansKR-Regular.otf", "NanumGothic.ttf", "UnDotum.ttf")
    return [os.path.join(root, name) for root in roots for name in names]


def _font_checks() -> List[Check]:
    paths = _font_paths()
    pretendard = next((p for p in paths if "Pretendard" in os.path.basename(p) and os.path.isfile(p)), None)
    fallback = next((p for p in paths if "Pretendard" not in os.path.basename(p) and os.path.isfile(p)), None)
    return [
        Check("font.pretendard", "PASS" if pretendard else "WARN",
              "Pretendard 글꼴을 찾았습니다." if pretendard else "Pretendard를 찾지 못했습니다.", pretendard),
        Check("font.fallback", "PASS" if fallback else "WARN",
              "플랫폼 대체 한글 글꼴을 찾았습니다." if fallback else "대체 한글 글꼴을 찾지 못했습니다.", fallback),
    ]


def _writer_check() -> Check:
    """Find a local PPTX writer without importing optional packages."""
    if importlib.util.find_spec("pptx") is not None:
        return Check("writer", "PASS", "Python python-pptx 제작 모듈을 찾았습니다.")
    node = shutil.which("node")
    npm = shutil.which("npm")
    if node and npm:
        try:
            result = subprocess.run([npm, "root", "-g"], capture_output=True, text=True,
                                    timeout=4, check=False)
            root = (result.stdout or "").strip()
            if root and os.path.isdir(os.path.join(root, "pptxgenjs")):
                return Check("writer", "PASS", "전역 pptxgenjs 제작 모듈을 찾았습니다.",
                             os.path.join(root, "pptxgenjs"))
        except (OSError, subprocess.SubprocessError):
            pass
    return Check("writer", "WARN", "별도 PPTX 제작 모듈을 찾지 못했습니다.")


def run_preflight(available_skills: Iterable[str] = ()) -> dict:
    system = platform.system() or "Unknown"
    checks: List[Check] = [
        Check("os", "PASS", f"{system} {platform.release()}"),
        Check("python", "PASS", platform.python_version(), sys.executable),
    ]
    command, path = _which(("powerpnt", "POWERPNT", "soffice", "libreoffice", "lowriter"))
    if not command:
        path = _installed_powerpoint()
        command = "POWERPNT" if path else None
    if command:
        version = "설치 확인" if command.upper() == "POWERPNT" else (_version(path or command) or "버전 확인 불가")
        checks.append(Check("renderer", "PASS", f"{command}: {version}", path))
    else:
        checks.append(Check("renderer", "WARN", "PowerPoint 또는 LibreOffice 계열 렌더러를 찾지 못했습니다."))
    writer = _writer_check()
    checks.append(writer)
    checks.extend(_font_checks())
    skills = {str(item).strip() for item in available_skills if str(item).strip()}
    found = "presentations:Presentations" in skills
    checks.append(Check("skill.presentations", "PASS" if found else "WARN",
                        "presentations:Presentations를 확인했습니다." if found else
                        "presentations:Presentations가 전달 목록에 없습니다."))
    tool_ok = command is not None or writer.status == "PASS"
    checks.append(Check("pptx.tooling", "PASS" if tool_ok else "BLOCKED",
                        "PPTX 제작 도구 또는 렌더러를 확인했습니다." if tool_ok else
                        "PPTX 제작 도구가 전혀 없어 제작을 진행할 수 없습니다."))
    blocked = not tool_ok
    warning_count = sum(c.status == "WARN" for c in checks)
    return {"mode": "compatibility", "blocked": blocked,
            "summary": "BLOCKED" if blocked else ("WARN" if warning_count else "PASS"),
            "platform": {"system": system, "release": platform.release(), "machine": platform.machine()},
            "checks": [asdict(c) for c in checks]}


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="cms-pptx-builder 호환 모드 환경 진단")
    parser.add_argument("--json", action="store_true", help="JSON으로 출력")
    parser.add_argument("--available-skill", action="append", default=[], help="skill 이름(반복 가능)")
    args = parser.parse_args(argv)
    result = run_preflight(args.available_skill)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"호환 모드 사전 진단: {result['summary']}")
        for check in result["checks"]:
            suffix = f" [{check['path']}]" if check.get("path") else ""
            print(f"- {check['status']}: {check['name']} — {check['detail']}{suffix}")
        print("시각 렌더링 완료 여부는 이 진단에서 판단하지 않습니다.")
    return 2 if result["blocked"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
