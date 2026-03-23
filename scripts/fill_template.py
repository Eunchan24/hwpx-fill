#!/usr/bin/env python3
"""
HWPX 양식 채우기 — 템플릿의 텍스트를 치환하여 최종 문서를 생성한다.

Usage:
    python3 fill_template.py <source.hwpx> <output.hwpx> <replacements.json> [--verify]

replacements.json 형식:
    {
        "single": {
            "브라더 공기관": "실제 기관명",
            "기본 보고서 양식": "실제 제목"
        },
        "sequential": {
            "헤드라인M 폰트 16포인트(문단 위 15)": ["값1", "값2", ...]
        }
    }
"""

import zipfile
import os
import re
import sys
import json
import shutil
from pathlib import Path


def zip_replace(src_path: str, dst_path: str, replacements: dict) -> int:
    """HWPX ZIP 내 모든 XML에서 텍스트 일괄 치환. 치환 횟수 반환."""
    total_replaced = 0
    tmp = dst_path + ".tmp"
    with zipfile.ZipFile(src_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                    text = data.decode("utf-8")
                    for old, new in replacements.items():
                        count = text.count(old)
                        if count > 0:
                            text = text.replace(old, new)
                            total_replaced += count
                    data = text.encode("utf-8")
                zout.writestr(item, data)
    if os.path.exists(dst_path):
        os.remove(dst_path)
    os.rename(tmp, dst_path)
    return total_replaced


def zip_replace_sequential(src_path: str, dst_path: str, old: str, new_list: list) -> int:
    """section XML에서 old를 순서대로 new_list 값으로 하나씩 치환. 치환 횟수 반환."""
    total_replaced = 0
    tmp = dst_path + ".tmp"
    with zipfile.ZipFile(src_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if "section" in item.filename and item.filename.endswith(".xml"):
                    text = data.decode("utf-8")
                    for new_val in new_list:
                        if old in text:
                            text = text.replace(old, new_val, 1)
                            total_replaced += 1
                    data = text.encode("utf-8")
                zout.writestr(item, data)
    if os.path.exists(dst_path):
        os.remove(dst_path)
    os.rename(tmp, dst_path)
    return total_replaced


def fix_namespaces(hwpx_path: str) -> None:
    """네임스페이스 프리픽스를 한컴오피스 표준으로 교체한다.

    python-hwpx가 생성한 ns0:/ns1: 등의 자동 프리픽스를
    한컴오피스가 요구하는 hh/hc/hp/hs로 바꾼다.
    이 처리를 하지 않으면 한글 뷰어에서 빈 페이지로 표시된다.
    """
    NS_MAP = {
        "http://www.hancom.co.kr/hwpml/2011/head": "hh",
        "http://www.hancom.co.kr/hwpml/2011/core": "hc",
        "http://www.hancom.co.kr/hwpml/2011/paragraph": "hp",
        "http://www.hancom.co.kr/hwpml/2011/section": "hs",
    }
    tmp = hwpx_path + ".tmp"
    with zipfile.ZipFile(hwpx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                    text = data.decode("utf-8")
                    ns_aliases = {}
                    for match in re.finditer(r'xmlns:(ns\d+)="([^"]+)"', text):
                        alias, uri = match.group(1), match.group(2)
                        if uri in NS_MAP:
                            ns_aliases[alias] = NS_MAP[uri]
                    for old_prefix, new_prefix in ns_aliases.items():
                        text = text.replace(f"xmlns:{old_prefix}=", f"xmlns:{new_prefix}=")
                        text = text.replace(f"<{old_prefix}:", f"<{new_prefix}:")
                        text = text.replace(f"</{old_prefix}:", f"</{new_prefix}:")
                    data = text.encode("utf-8")
                zout.writestr(item, data)
    os.replace(tmp, hwpx_path)


def fill_template(src: str, dst: str, single: dict, sequential: dict) -> dict:
    """템플릿 복사 -> 단일 치환 -> 순차 치환 -> 네임스페이스 후처리.

    Returns:
        치환 통계 dict
    """
    shutil.copy2(src, dst)
    stats = {"single_replaced": 0, "sequential_replaced": 0}

    if single:
        stats["single_replaced"] = zip_replace(dst, dst, single)

    for old_text, new_list in sequential.items():
        stats["sequential_replaced"] += zip_replace_sequential(dst, dst, old_text, new_list)

    fix_namespaces(dst)
    return stats


def verify_output(hwpx_path: str) -> list:
    """치환 결과를 검증하여 남은 텍스트 목록 반환."""
    from hwpx import ObjectFinder

    finder = ObjectFinder(hwpx_path)
    results = finder.find_all(tag="t")
    texts = []
    for r in results:
        if r.text and r.text.strip():
            texts.append(r.text)
    return texts


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python3 fill_template.py <source.hwpx> <output.hwpx> <replacements.json> [--verify]")
        print()
        print("replacements.json format:")
        print('  {"single": {"old": "new", ...}, "sequential": {"old": ["v1","v2",...], ...}}')
        sys.exit(1)

    src = sys.argv[1]
    dst = sys.argv[2]
    rep_file = sys.argv[3]

    if not Path(src).exists():
        print(f"Error: Source not found: {src}", file=sys.stderr)
        sys.exit(1)

    with open(rep_file, "r", encoding="utf-8") as f:
        replacements = json.load(f)

    single = replacements.get("single", {})
    sequential = replacements.get("sequential", {})

    stats = fill_template(src, dst, single, sequential)
    print(f"완료: {dst}")
    print(f"  단일 치환: {stats['single_replaced']}건")
    print(f"  순차 치환: {stats['sequential_replaced']}건")

    if "--verify" in sys.argv:
        remaining = verify_output(dst)
        print(f"  남은 텍스트: {len(remaining)}개")
        for t in remaining:
            print(f"    {repr(t)}")
