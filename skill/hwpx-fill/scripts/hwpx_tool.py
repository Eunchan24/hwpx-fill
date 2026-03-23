#!/usr/bin/env python3
"""
HWPX Tool — HWPX 양식 분석, 치환, 미리보기, 네임스페이스 수정 통합 도구.

Subcommands:
    analyze  <file>                      — 양식 텍스트 추출 + 분류 (JSON)
    fill     <src> <dst> <replacements>  — 단일 패스 치환 + NS 수정
    preview  <file>                      — Markdown 미리보기 생성
    fix      <file>                      — 네임스페이스만 단독 수정

Usage:
    python hwpx_tool.py analyze template.hwpx
    python hwpx_tool.py fill template.hwpx output.hwpx replacements.json
    python hwpx_tool.py preview output.hwpx
    python hwpx_tool.py fix output.hwpx

replacements.json format:
    {
        "single": {"old_text": "new_text", ...},
        "sequential": {"repeated_text": ["val1", "val2", ...]}
    }
"""

import argparse
import collections
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


# ─── Constants ─────────────────────────────────────────────────────────────

SKILL_DIR = Path(__file__).resolve().parent.parent
ASSETS_DIR = SKILL_DIR / "assets"

NS_MAP = {
    "http://www.hancom.co.kr/hwpml/2011/head": "hh",
    "http://www.hancom.co.kr/hwpml/2011/core": "hc",
    "http://www.hancom.co.kr/hwpml/2011/paragraph": "hp",
    "http://www.hancom.co.kr/hwpml/2011/section": "hs",
}


# ─── Dependency Check ──────────────────────────────────────────────────────

def _ensure_hwpx():
    """python-hwpx 라이브러리 확인. 없으면 자동 설치 시도."""
    try:
        import hwpx  # noqa: F401
        return True
    except ImportError:
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "python-hwpx",
                 "--break-system-packages", "-q"],
                check=True, capture_output=True,
            )
            return True
        except subprocess.CalledProcessError:
            _die("python-hwpx 설치 실패. 수동 설치 필요: pip install python-hwpx")


def _die(msg):
    """에러 JSON 출력 후 종료."""
    print(json.dumps({"error": msg}, ensure_ascii=False))
    sys.exit(1)


# ─── XML Processing: Run 병합 (쪼개진 텍스트 복원) ─────────────────────────

def _normalize_split_runs(text):
    """
    동일 서식(charPrIDRef)의 인접 <run> 태그를 병합하여 쪼개진 텍스트를 복원.

    Before:
      <hp:run><hp:rPr charPrIDRef="5"/><hp:t>브라더 </hp:t></hp:run>
      <hp:run><hp:rPr charPrIDRef="5"/><hp:t>공기관</hp:t></hp:run>
    After:
      <hp:run><hp:rPr charPrIDRef="5"/><hp:t>브라더 공기관</hp:t></hp:run>

    동일 charPrIDRef인 경우만 병합하므로 서식이 다른 run은 유지된다.
    """
    # 인접 run 경계 패턴:
    #   </ns:t> </ns:run> <ns:run> <ns:rPr charPrIDRef="N" .../> <ns:t>
    boundary = re.compile(
        r'</(\w+):t>'                                 # close </ns:t>
        r'\s*'
        r'</\1:run>'                                  # close </ns:run>
        r'\s*'
        r'<\1:run>'                                   # open <ns:run>
        r'\s*'
        r'<\1:rPr\s+charPrIDRef="(\d+)"[^/]*/>'      # <ns:rPr charPrIDRef="N".../>
        r'\s*'
        r'<\1:t>'                                     # open <ns:t>
    )

    changed = True
    while changed:
        changed = False
        for m in boundary.finditer(text):
            ns, ref_id = m.group(1), m.group(2)
            # 앞쪽 run의 charPrIDRef가 동일한지 확인
            before = text[:m.start()]
            prev_pat = (
                rf'<{re.escape(ns)}:rPr\s+charPrIDRef="{re.escape(ref_id)}"'
                rf'[^/]*/>\s*<{re.escape(ns)}:t>[^<]*$'
            )
            if re.search(prev_pat, before):
                # 동일 서식 → 경계 제거 (run 병합)
                text = text[:m.start()] + text[m.end():]
                changed = True
                break

    return text


def _try_split_replace(xml_text, old, new):
    """
    <run> 경계를 넘어 쪼개진 텍스트도 치환 시도.

    Simple replace 실패 시, <t> 태그 경계 사이에 분산된 텍스트를
    regex로 찾아 치환한다.

    Returns: (modified_text, success: bool)
    """
    # 1차: 단순 치환
    if old in xml_text:
        return xml_text.replace(old, new, 1), True

    # 2차: run 경계를 허용하는 패턴으로 검색
    # 허용하는 경계: </ns:t></ns:run><ns:run><ns:rPr .../><ns:t>
    run_boundary = (
        r'(?:'
        r'</\w+:t>\s*</\w+:run>\s*<\w+:run>\s*<\w+:rPr[^/]*/>\s*<\w+:t>'
        r')?'
    )

    parts = [re.escape(old[0])]
    for ch in old[1:]:
        parts.append(run_boundary)
        parts.append(re.escape(ch))

    pattern = ''.join(parts)
    match = re.search(pattern, xml_text)

    if match:
        # 태그 제거 후 원본 텍스트와 일치하는지 검증
        stripped = re.sub(r'<[^>]+>', '', match.group(0))
        if stripped == old:
            # 매칭 영역을 새 텍스트로 교체
            # (중간 run 태그가 제거되어 하나의 <t> 안에 합쳐짐)
            xml_text = xml_text[:match.start()] + new + xml_text[match.end():]
            return xml_text, True

    return xml_text, False


# ─── XML Processing: 네임스페이스 교정 ─────────────────────────────────────

def _fix_ns_in_text(text):
    """XML 텍스트의 ns0/ns1 → hh/hc/hp/hs 네임스페이스 프리픽스 교정."""
    ns_aliases = {}
    for m in re.finditer(r'xmlns:(ns\d+)="([^"]+)"', text):
        alias, uri = m.group(1), m.group(2)
        if uri in NS_MAP:
            ns_aliases[alias] = NS_MAP[uri]

    for old_p, new_p in ns_aliases.items():
        text = text.replace(f"xmlns:{old_p}=", f"xmlns:{new_p}=")
        text = text.replace(f"<{old_p}:", f"<{new_p}:")
        text = text.replace(f"</{old_p}:", f"</{new_p}:")

    return text


# ─── Core: analyze ─────────────────────────────────────────────────────────

def analyze(hwpx_path):
    """
    양식 분석: <t> 태그 텍스트 추출 + 단일/순차 분류 + 문단 구조.

    Returns: dict (JSON 직렬화 가능)
    """
    _ensure_hwpx()
    from hwpx import ObjectFinder, TextExtractor

    # <t> 태그 원문 추출
    finder = ObjectFinder(hwpx_path)
    raw_results = finder.find_all(tag="t")
    raw_texts = [r.text for r in raw_results if r.text and r.text.strip()]

    # 문단 구조 추출
    paragraphs = []
    ext = TextExtractor(hwpx_path)
    ext.open()
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

    # 중복 카운트 + 분류
    counter = collections.Counter(raw_texts)
    entries = []
    seen = set()
    for t in raw_texts:
        if t not in seen:
            seen.add(t)
            count = counter[t]
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
    }


# ─── Core: fill (단일 패스) ────────────────────────────────────────────────

def fill(src_path, dst_path, single, sequential):
    """
    단일 패스 치환: ZIP 1회 읽기 → run 병합 + 치환 + NS 수정 → 1회 쓰기.

    6회 ZIP 사이클 → 1회로 최적화.

    Args:
        src_path:   원본 HWPX 경로
        dst_path:   출력 HWPX 경로
        single:     {"old": "new", ...} 일괄 치환
        sequential: {"old": ["v1", "v2", ...], ...} 순차 치환

    Returns:
        통계 dict
    """
    stats = {
        "single_replaced": 0,
        "sequential_replaced": 0,
        "split_fixed": 0,
    }
    seq_queues = {old: list(vals) for old, vals in sequential.items()}

    tmp = dst_path + ".tmp"
    with zipfile.ZipFile(src_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                    text = data.decode("utf-8")

                    # ① 쪼개진 run 병합 (전처리)
                    pre_len = len(text)
                    text = _normalize_split_runs(text)
                    if len(text) != pre_len:
                        stats["split_fixed"] += 1

                    # ② 단일 치환
                    for old, new in single.items():
                        count = text.count(old)
                        if count > 0:
                            text = text.replace(old, new)
                            stats["single_replaced"] += count
                        else:
                            # run 병합 후에도 못 찾으면 cross-run 치환 시도
                            text, ok = _try_split_replace(text, old, new)
                            if ok:
                                stats["single_replaced"] += 1
                                stats["split_fixed"] += 1

                    # ③ 순차 치환 (section 파일만)
                    if "section" in item.filename:
                        for old, queue in seq_queues.items():
                            # 일반 순차 치환
                            while queue and old in text:
                                text = text.replace(old, queue.pop(0), 1)
                                stats["sequential_replaced"] += 1
                            # cross-run 순차 치환
                            while queue:
                                text, ok = _try_split_replace(text, old, queue[0])
                                if ok:
                                    queue.pop(0)
                                    stats["sequential_replaced"] += 1
                                    stats["split_fixed"] += 1
                                else:
                                    break

                    # ④ 네임스페이스 수정
                    text = _fix_ns_in_text(text)

                    data = text.encode("utf-8")

                zout.writestr(item, data)

    if os.path.exists(dst_path):
        os.remove(dst_path)
    os.rename(tmp, dst_path)

    return stats


# ─── Core: preview ─────────────────────────────────────────────────────────

def preview(hwpx_path):
    """
    HWPX 파일의 Markdown 미리보기 생성.

    문단 텍스트를 추출하고 □○―※ 계층과 섹션 구조를 반영한
    Markdown을 반환한다. Cloud Claude 채팅에서 바로 확인 가능.
    """
    _ensure_hwpx()
    from hwpx import TextExtractor

    ext = TextExtractor(hwpx_path)
    ext.open()

    lines = []

    for sec_idx, sec in enumerate(ext.iter_sections()):
        lines.append(f"\n---\n## 섹션 {sec_idx + 1}\n")

        for para in ext.iter_paragraphs(sec):
            text = para.text()
            if not text.strip():
                continue

            s = text.strip()

            # 보고서 기호 체계 계층화
            if s.startswith("□"):
                lines.append(f"\n**{s}**")
            elif s.startswith("○"):
                lines.append(f"&ensp;&ensp;{s}")
            elif s.startswith("―"):
                lines.append(f"&ensp;&ensp;&ensp;&ensp;{s}")
            elif s.startswith("※"):
                lines.append(f"&ensp;&ensp;&ensp;&ensp;&ensp;&ensp;*{s}*")
            # 로마숫자 섹션 제목
            elif s and s[0] in "ⅠⅡⅢⅣⅤ":
                lines.append(f"\n### {s}")
            # 공문서 번호 계층
            elif re.match(r'^\d+\.', s):
                lines.append(f"\n**{s}**")
            elif re.match(r'^[가-힣]\.', s):
                lines.append(f"&ensp;&ensp;{s}")
            elif re.match(r'^\d+\)', s):
                lines.append(f"&ensp;&ensp;&ensp;&ensp;{s}")
            else:
                lines.append(s)

    ext.close()
    return "\n".join(lines)


# ─── Core: fix ─────────────────────────────────────────────────────────────

def fix_namespaces(hwpx_path):
    """네임스페이스 프리픽스만 단독 수정. doc.save() 후처리용."""
    tmp = hwpx_path + ".tmp"
    fixed = 0

    with zipfile.ZipFile(hwpx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                    text = data.decode("utf-8")
                    new_text = _fix_ns_in_text(text)
                    if new_text != text:
                        fixed += 1
                    data = new_text.encode("utf-8")
                zout.writestr(item, data)

    os.replace(tmp, hwpx_path)
    return fixed


# ─── Core: verify ──────────────────────────────────────────────────────────

def verify(hwpx_path):
    """치환 결과 검증: 남은 텍스트 목록 반환."""
    _ensure_hwpx()
    from hwpx import ObjectFinder

    finder = ObjectFinder(hwpx_path)
    results = finder.find_all(tag="t")
    return [r.text for r in results if r.text and r.text.strip()]


# ─── CLI ───────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="HWPX Tool — analyze, fill, preview, fix",
    )
    sub = parser.add_subparsers(dest="command")

    # analyze
    p = sub.add_parser("analyze", help="양식 텍스트 분석 (JSON 출력)")
    p.add_argument("file", help=".hwpx 파일 경로")

    # fill
    p = sub.add_parser("fill", help="단일 패스 치환 + NS 수정")
    p.add_argument("src", help="원본 .hwpx")
    p.add_argument("dst", help="출력 .hwpx")
    p.add_argument("replacements", help="replacements.json 경로")

    # preview
    p = sub.add_parser("preview", help="Markdown 미리보기 생성")
    p.add_argument("file", help=".hwpx 파일 경로")

    # fix
    p = sub.add_parser("fix", help="네임스페이스만 단독 수정")
    p.add_argument("file", help=".hwpx 파일 경로")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # ── analyze ──
    if args.command == "analyze":
        if not Path(args.file).exists():
            _die(f"파일 없음: {args.file}")
        result = analyze(args.file)
        print(json.dumps(result, ensure_ascii=False, indent=2))

    # ── fill ──
    elif args.command == "fill":
        if not Path(args.src).exists():
            _die(f"원본 파일 없음: {args.src}")
        if not Path(args.replacements).exists():
            _die(f"치환 파일 없음: {args.replacements}")

        with open(args.replacements, "r", encoding="utf-8") as f:
            rep = json.load(f)

        single = rep.get("single", {})
        sequential = rep.get("sequential", {})

        stats = fill(args.src, args.dst, single, sequential)

        # 검증
        remaining = verify(args.dst)

        print(json.dumps({
            "status": "success",
            "output": str(Path(args.dst).resolve()),
            "stats": stats,
            "remaining_count": len(remaining),
            "remaining_sample": remaining[:20],
        }, ensure_ascii=False, indent=2))

    # ── preview ──
    elif args.command == "preview":
        if not Path(args.file).exists():
            _die(f"파일 없음: {args.file}")
        print(preview(args.file))

    # ── fix ──
    elif args.command == "fix":
        if not Path(args.file).exists():
            _die(f"파일 없음: {args.file}")
        fixed = fix_namespaces(args.file)
        print(json.dumps({
            "status": "success",
            "fixed_files": fixed,
            "path": str(Path(args.file).resolve()),
        }, ensure_ascii=False))


if __name__ == "__main__":
    main()
