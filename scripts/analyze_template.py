#!/usr/bin/env python3
"""
HWPX 양식 분석기 — 템플릿 내 치환 가능한 텍스트를 추출하고 분류한다.

Usage:
    python3 analyze_template.py <template.hwpx> [--format json|md]

Output:
    JSON (default) 또는 Markdown 형식의 분석 결과
"""

import sys
import json
import collections
from pathlib import Path


def analyze_template(hwpx_path: str) -> dict:
    """HWPX 템플릿을 분석하여 치환 가능한 텍스트 목록을 반환한다."""
    from hwpx import ObjectFinder, TextExtractor

    # 1단계: ObjectFinder로 <t> 태그 원문 추출 (치환 타겟용)
    finder = ObjectFinder(hwpx_path)
    raw_results = finder.find_all(tag="t")

    raw_texts = []
    for r in raw_results:
        if r.text and r.text.strip():
            raw_texts.append(r.text)

    # 2단계: TextExtractor로 문단 단위 텍스트 추출 (구조 파악용)
    ext = TextExtractor(hwpx_path)
    ext.open()
    paragraphs = []
    for sec_idx, sec in enumerate(ext.iter_sections()):
        for para_idx, para in enumerate(ext.iter_paragraphs(sec)):
            text = para.text()
            if text.strip():
                paragraphs.append({
                    "section": sec_idx,
                    "index": para_idx,
                    "text": text,
                })
    ext.close()

    # 3단계: 중복 카운트 및 치환 타입 분류
    text_counter = collections.Counter(raw_texts)

    entries = []
    seen = set()
    for t in raw_texts:
        if t not in seen:
            seen.add(t)
            count = text_counter[t]
            entries.append({
                "text": t,
                "count": count,
                "type": "sequential" if count > 1 else "single",
            })

    return {
        "path": str(hwpx_path),
        "total_text_elements": len(raw_texts),
        "unique_texts": len(entries),
        "entries": entries,
        "paragraphs": paragraphs,
        "raw_order": raw_texts,
    }


def format_markdown(analysis: dict) -> str:
    """분석 결과를 사용자 친화적 Markdown으로 변환한다."""
    lines = []
    lines.append("## HWPX 양식 분석 결과\n")
    lines.append(f"- **파일**: `{analysis['path']}`")
    lines.append(f"- **텍스트 요소**: {analysis['total_text_elements']}개")
    lines.append(f"- **고유 텍스트**: {analysis['unique_texts']}개\n")

    # 단일 치환 항목
    singles = [e for e in analysis["entries"] if e["type"] == "single"]
    if singles:
        lines.append("### 단일 치환 항목\n")
        for i, e in enumerate(singles, 1):
            lines.append(f'{i}. `{e["text"]}`')
        lines.append("")

    # 순차 치환 항목
    sequentials = [e for e in analysis["entries"] if e["type"] == "sequential"]
    if sequentials:
        lines.append("### 순차 치환 항목 (동일 텍스트 반복)\n")
        for e in sequentials:
            lines.append(f'- `{e["text"]}` x **{e["count"]}개**')
        lines.append("")

    # 문서 구조 (문단 순서)
    lines.append("### 문서 구조 (문단 순서)\n")
    current_section = -1
    for p in analysis["paragraphs"]:
        if p["section"] != current_section:
            current_section = p["section"]
            lines.append(f"\n**섹션 {current_section}**\n")
        lines.append(f'  `[{p["index"]}]` {p["text"]}')

    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 analyze_template.py <template.hwpx> [--format json|md]")
        sys.exit(1)

    path = sys.argv[1]
    fmt = "json"
    if "--format" in sys.argv:
        idx = sys.argv.index("--format")
        if idx + 1 < len(sys.argv):
            fmt = sys.argv[idx + 1]

    if not Path(path).exists():
        print(f"Error: File not found: {path}", file=sys.stderr)
        sys.exit(1)

    analysis = analyze_template(path)

    if fmt == "md":
        print(format_markdown(analysis))
    else:
        print(json.dumps(analysis, ensure_ascii=False, indent=2))
