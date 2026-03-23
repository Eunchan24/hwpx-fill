#!/usr/bin/env python3
"""
HWPX 양식 채우기 — 템플릿의 텍스트를 치환하여 최종 문서를 생성한다.

원본 ZIP 바이너리를 최대한 보존한다:
- 비수정 엔트리: 원본 compressed bytes 그대로 복사
- 수정 엔트리: 텍스트 치환 후 동일 압축 방식으로 재압축

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

import os
import re
import struct
import sys
import json
import zlib
import shutil
import zipfile
from pathlib import Path


# --- Low-level ZIP manipulation ---

# ZIP local file header signature
_LFH_SIG = b"PK\x03\x04"
_LFH_SIZE = 30  # fixed part

# ZIP central directory header signature
_CDH_SIG = b"PK\x01\x02"
_CDH_SIZE = 46  # fixed part

# End of central directory signature
_EOCD_SIG = b"PK\x05\x06"
_EOCD_SIZE = 22  # fixed part (without comment)


def _read_eocd(data: bytes) -> dict:
    """ZIP End of Central Directory 레코드 파싱."""
    # EOCD는 파일 끝에서 검색 (코멘트 가능하므로 뒤에서부터)
    pos = data.rfind(_EOCD_SIG)
    if pos < 0:
        raise ValueError("EOCD not found")
    fields = struct.unpack_from("<4sHHHHIIH", data, pos)
    return {
        "offset": pos,
        "disk_num": fields[1],
        "cd_disk_num": fields[2],
        "cd_entries_this_disk": fields[3],
        "cd_entries_total": fields[4],
        "cd_size": fields[5],
        "cd_offset": fields[6],
        "comment_len": fields[7],
    }


def _read_central_directory(data: bytes, eocd: dict) -> list:
    """Central Directory의 모든 엔트리를 파싱."""
    entries = []
    pos = eocd["cd_offset"]
    for _ in range(eocd["cd_entries_total"]):
        if data[pos:pos+4] != _CDH_SIG:
            raise ValueError(f"Invalid CDH at offset {pos}")
        (sig, ver_made, ver_need, flags, method,
         mod_time, mod_date, crc, comp_size, uncomp_size,
         fn_len, extra_len, comment_len,
         disk_start, internal_attr, external_attr,
         local_offset) = struct.unpack_from("<4sHHHHHHIIIHHHHHII", data, pos)
        filename = data[pos+_CDH_SIZE:pos+_CDH_SIZE+fn_len].decode("utf-8")
        entry = {
            "cd_offset": pos,
            "version_made_by": ver_made,
            "version_needed": ver_need,
            "flags": flags,
            "compression": method,
            "mod_time": mod_time,
            "mod_date": mod_date,
            "crc32": crc,
            "compressed_size": comp_size,
            "uncompressed_size": uncomp_size,
            "filename": filename,
            "fn_len": fn_len,
            "extra_len": extra_len,
            "comment_len": comment_len,
            "disk_start": disk_start,
            "internal_attr": internal_attr,
            "external_attr": external_attr,
            "local_header_offset": local_offset,
            "extra": data[pos+_CDH_SIZE+fn_len:pos+_CDH_SIZE+fn_len+extra_len],
            "comment": data[pos+_CDH_SIZE+fn_len+extra_len:pos+_CDH_SIZE+fn_len+extra_len+comment_len],
        }
        entries.append(entry)
        pos += _CDH_SIZE + fn_len + extra_len + comment_len
    return entries


def _read_local_file_data(data: bytes, entry: dict) -> bytes:
    """Local file header에서 compressed data를 raw bytes로 읽기."""
    off = entry["local_header_offset"]
    if data[off:off+4] != _LFH_SIG:
        raise ValueError(f"Invalid LFH at offset {off}")
    # Local header에서 fn_len, extra_len 다시 읽기 (CD와 다를 수 있음)
    lfh_fn_len, lfh_extra_len = struct.unpack_from("<HH", data, off + 26)
    data_start = off + _LFH_SIZE + lfh_fn_len + lfh_extra_len
    return data[data_start:data_start + entry["compressed_size"]]


def _build_local_file_header(entry: dict) -> bytes:
    """Local file header를 빌드."""
    fn_bytes = entry["filename"].encode("utf-8")
    extra = entry.get("local_extra", entry["extra"])
    header = struct.pack(
        "<4sHHHHHIIIHH",
        _LFH_SIG,
        entry["version_needed"],
        entry["flags"],
        entry["compression"],
        entry["mod_time"],
        entry["mod_date"],
        entry["crc32"],
        entry["compressed_size"],
        entry["uncompressed_size"],
        len(fn_bytes),
        len(extra),
    )
    return header + fn_bytes + extra


def _build_central_directory_header(entry: dict) -> bytes:
    """Central Directory header를 빌드."""
    fn_bytes = entry["filename"].encode("utf-8")
    header = struct.pack(
        "<4sHHHHHHIIIHHHHHII",
        _CDH_SIG,
        entry["version_made_by"],
        entry["version_needed"],
        entry["flags"],
        entry["compression"],
        entry["mod_time"],
        entry["mod_date"],
        entry["crc32"],
        entry["compressed_size"],
        entry["uncompressed_size"],
        len(fn_bytes),
        len(entry["extra"]),
        len(entry["comment"]),
        entry["disk_start"],
        entry["internal_attr"],
        entry["external_attr"],
        entry["local_header_offset"],
    )
    return header + fn_bytes + entry["extra"] + entry["comment"]


def raw_zip_replace(src_path: str, dst_path: str, modifications: dict) -> None:
    """원본 ZIP의 바이너리를 최대한 보존하면서 지정된 엔트리만 교체.

    Args:
        src_path: 원본 HWPX 경로
        dst_path: 출력 HWPX 경로
        modifications: {filename: new_content_bytes} 딕셔너리
    """
    with open(src_path, "rb") as f:
        src_data = f.read()

    eocd = _read_eocd(src_data)
    entries = _read_central_directory(src_data, eocd)

    with open(dst_path, "wb") as out:
        new_entries = []

        for entry in entries:
            fn = entry["filename"]
            new_offset = out.tell()

            if fn in modifications:
                # 수정된 엔트리: 새 데이터로 압축
                new_data = modifications[fn]
                new_crc = zlib.crc32(new_data) & 0xFFFFFFFF
                # 원본과 동일한 압축 방식 사용
                if entry["compression"] == 8:  # DEFLATED
                    compressed = zlib.compress(new_data, 1)  # level 1 = fastest
                    # zlib.compress adds a 2-byte header and 4-byte checksum
                    # ZIP DEFLATE uses raw deflate without header/checksum
                    compressed_raw = compressed[2:-4]
                elif entry["compression"] == 0:  # STORED
                    compressed_raw = new_data
                else:
                    raise ValueError(f"Unsupported compression: {entry['compression']}")

                # 엔트리 정보 업데이트
                new_entry = dict(entry)
                new_entry["crc32"] = new_crc
                new_entry["compressed_size"] = len(compressed_raw)
                new_entry["uncompressed_size"] = len(new_data)
                new_entry["local_header_offset"] = new_offset

                # LFH의 extra는 원본 LFH에서 읽기
                off = entry["local_header_offset"]
                lfh_fn_len, lfh_extra_len = struct.unpack_from("<HH", src_data, off + 26)
                new_entry["local_extra"] = src_data[
                    off + _LFH_SIZE + lfh_fn_len:
                    off + _LFH_SIZE + lfh_fn_len + lfh_extra_len
                ]

                lfh = _build_local_file_header(new_entry)
                out.write(lfh)
                out.write(compressed_raw)
                new_entries.append(new_entry)
            else:
                # 비수정 엔트리: 원본 raw bytes 그대로 복사
                off = entry["local_header_offset"]
                lfh_fn_len, lfh_extra_len = struct.unpack_from("<HH", src_data, off + 26)
                total_len = (
                    _LFH_SIZE + lfh_fn_len + lfh_extra_len + entry["compressed_size"]
                )
                raw = src_data[off:off + total_len]

                new_entry = dict(entry)
                new_entry["local_header_offset"] = new_offset
                out.write(raw)
                new_entries.append(new_entry)

        # Central Directory 작성
        cd_offset = out.tell()
        cd_size = 0
        for entry in new_entries:
            cdh = _build_central_directory_header(entry)
            out.write(cdh)
            cd_size += len(cdh)

        # End of Central Directory 작성
        eocd_record = struct.pack(
            "<4sHHHHIIH",
            _EOCD_SIG,
            0,  # disk number
            0,  # CD start disk
            len(new_entries),  # entries this disk
            len(new_entries),  # entries total
            cd_size,
            cd_offset,
            0,  # comment length
        )
        out.write(eocd_record)


# --- High-level template filling ---


def _fix_namespaces_text(text: str) -> str:
    """XML 텍스트의 네임스페이스 프리픽스를 한컴오피스 표준으로 교체.

    ns0:/ns1: 등의 자동 프리픽스를 hh/hc/hp/hs로 바꾼다.
    """
    NS_MAP = {
        "http://www.hancom.co.kr/hwpml/2011/head": "hh",
        "http://www.hancom.co.kr/hwpml/2011/core": "hc",
        "http://www.hancom.co.kr/hwpml/2011/paragraph": "hp",
        "http://www.hancom.co.kr/hwpml/2011/section": "hs",
    }

    ns_aliases = {}
    for match in re.finditer(r'xmlns:(ns\d+)="([^"]+)"', text):
        alias, uri = match.group(1), match.group(2)
        if uri in NS_MAP:
            ns_aliases[alias] = NS_MAP[uri]

    if not ns_aliases:
        return text

    for old_prefix, new_prefix in ns_aliases.items():
        text = text.replace(f"xmlns:{old_prefix}=", f"xmlns:{new_prefix}=")
        text = text.replace(f"<{old_prefix}:", f"<{new_prefix}:")
        text = text.replace(f"</{old_prefix}:", f"</{new_prefix}:")

    return text


def fill_template(src: str, dst: str, single: dict, sequential: dict) -> dict:
    """템플릿을 채워서 최종 문서를 생성한다.

    1단계: 원본 ZIP에서 수정 대상 XML을 읽고 텍스트 치환
    2단계: raw_zip_replace로 수정된 엔트리만 교체 (나머지는 원본 bytes 보존)

    Returns:
        치환 통계 dict
    """
    stats = {"single_replaced": 0, "sequential_replaced": 0}
    modifications = {}  # {filename: new_content_bytes}

    # 1. 원본에서 XML 텍스트 추출 & 치환
    with zipfile.ZipFile(src, "r") as zf:
        for item in zf.infolist():
            fn = item.filename
            if not (fn.startswith("Contents/") and fn.endswith(".xml")):
                continue

            text = zf.read(fn).decode("utf-8")
            original_text = text

            # 일괄 치환
            if single:
                for old, new in single.items():
                    n = text.count(old)
                    if n > 0:
                        text = text.replace(old, new)
                        stats["single_replaced"] += n

            # 순차 치환 (section 파일만)
            if "section" in fn:
                for old_text, new_list in sequential.items():
                    for new_val in new_list:
                        if old_text in text:
                            text = text.replace(old_text, new_val, 1)
                            stats["sequential_replaced"] += 1

            # 네임스페이스 후처리
            text = _fix_namespaces_text(text)

            # 변경된 경우만 수정 대상에 추가
            if text != original_text:
                modifications[fn] = text.encode("utf-8")

    # 2. 원본 ZIP을 복사하되, 수정된 엔트리만 교체
    if modifications:
        raw_zip_replace(src, dst, modifications)
        print(f"  수정된 파일: {', '.join(modifications.keys())}")
    else:
        shutil.copy2(src, dst)
        print("  변경 없음 — 원본 복사")

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
