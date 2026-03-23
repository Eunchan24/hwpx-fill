---
name: hwpx-fill
description: "HWPX 한글 문서(.hwpx)를 생성, 편집, 템플릿 치환하는 스킬. '한글 문서', 'hwpx', '보고서 양식', '공문서', '기안문', '한글로 작성', '양식 채워줘', 'hwpx fill' 등의 키워드에 사용. python-hwpx 라이브러리로 ZIP+XML 구조를 직접 처리하며 원본 서식을 100% 보존한다. Word(.docx)에는 docx 스킬을 사용할 것."
compatibility: "Python 3, pip install python-hwpx"
metadata:
  author: eunchankim
  version: "2.0"
---

# HWPX 문서 생성/편집 스킬

HWPX는 한컴오피스 한글의 개방형 문서 포맷(ZIP + XML, KS X 6101/OWPML 표준). 이 스킬은 `scripts/hwpx_tool.py` 통합 도구로 양식 분석, 치환, 미리보기, 네임스페이스 수정을 수행한다.

## 설치

```bash
pip install python-hwpx --break-system-packages
```

> `hwpx_tool.py`가 미설치 감지 시 자동 설치를 시도한다.

---

## 양식(템플릿) 선택 정책 — 예외 없음

### 1순위: 사용자 업로드 양식

사용자가 `.hwpx` 파일을 업로드했다면 **반드시** 해당 파일을 템플릿으로 사용한다.
- Claude.ai: `/mnt/user-data/uploads/` 에서 `.hwpx` 확인
- Claude Code: 사용자가 제공한 로컬 경로 사용

### 2순위: 기본 제공 양식

업로드 없으면 `assets/report-template.hwpx` 사용.

### 3순위: HwpxDocument.new() — 최후의 수단

아주 단순한 메모/목록 수준의 문서에만 허용. 보고서, 공문, 기안문은 절대 new()로 만들지 않는다.

---

## 4단계 워크플로우

### Phase 1: 양식 확보 + 분석

양식 파일을 작업 디렉토리로 복사한 후 분석:

```bash
python scripts/hwpx_tool.py analyze <template.hwpx>
```

JSON 출력에서 `entries`의 `type`을 확인:
- `"single"`: 일괄 치환 대상 (1회 등장)
- `"sequential"`: 순차 치환 대상 (N회 반복)

### Phase 2: 인터랙티브 값 수집

분석 결과를 **위치별 그룹핑 표**로 사용자에게 제시:

```
[표지] 기관명, 제목, 작성일
[본문] □ 항목 8개, ○ 항목 8개, ― 항목 8개, ※ 항목 7개
```

값 수집 규칙:
- 자연어 답변을 파싱하여 매핑 ("기관명은 팀워크" → 자동 매핑)
- 부분 입력 허용 ("표지만 먼저 채워줘")
- "AI가 알아서 채워줘" → 주제 기반 자동 생성
- 빈 값은 원본 유지
- 순차 항목은 **순서대로 N개** 필요함을 명확히 안내

### Phase 3: 확인

치환 매핑을 단일/순차 구분하여 전체 표로 보여준다. **사용자 승인 후에만** 실행. 수정 요청 시 Phase 2로 복귀.

### Phase 4: 실행 + 미리보기

1. `replacements.json` 생성:
```json
{
  "single": {"브라더 공기관": "실제 기관명", ...},
  "sequential": {"헤드라인M 폰트 16포인트(문단 위 15)": ["값1", "값2", ...]}
}
```

2. 치환 실행 (단일 패스 — ZIP 1회 해제/압축):
```bash
python scripts/hwpx_tool.py fill <src.hwpx> <dst.hwpx> replacements.json
```

3. Markdown 미리보기 표시:
```bash
python scripts/hwpx_tool.py preview <dst.hwpx>
```

4. 사용자 확인 후 파일 전달:
   - Claude.ai: `/mnt/user-data/outputs/` 복사 → `present_files`
   - Claude Code: 사용자 지정 경로 또는 `output/`

---

## 기본 양식(report-template.hwpx) 플레이스홀더

| 플레이스홀더 | 위치 | 치환 방법 |
|------------|------|----------|
| `브라더 공기관` | 표지 기관명 | 일괄 |
| `기본 보고서 양식` | 표지 제목 | 일괄 |
| `2024. 5. 23.` | 표지 작성일 | 일괄 |
| `제 목` | 본문 제목 | 일괄 |
| `. 개요` 등 | 목차 항목 | 일괄 |
| `추진 배경` 등 | 섹션 바 제목 | 일괄 |
| `헤드라인M 폰트 16포인트(문단 위 15)` | □ 본문 x8 | **순차** |
| `○ 휴면명조 15포인트(문단위 10)` | ○ 본문 x8 | **순차** |
| `― 휴면명조 15포인트(문단 위 6)` | ― 본문 x8 | **순차** |
| `※ 중고딕 13포인트(문단 위 3)` | ※ 주석 x7 | **순차** |

### 본문 기호 체계 (보고서 전용)

```
□  (HY헤드라인M 16pt) — 1단계
 ○  (휴먼명조 15pt)   — 2단계
  ―  (휴먼명조 15pt)  — 3단계
   ※  (한양중고딕 13pt) — 4단계
```

---

## 문서 유형별 스타일 가이드

| 유형 | 참조 파일 |
|------|----------|
| 보고서 (내부 보고용) | `references/report-style.md` |
| 공문서 (기안문) | `references/official-doc-style.md` |
| 저수준 XML 조작 | `references/xml-internals.md` |

---

## hwpx_tool.py 서브커맨드 요약

| 커맨드 | 용도 | 출력 |
|--------|------|------|
| `analyze <file>` | 양식 텍스트 분석 | JSON |
| `fill <src> <dst> <json>` | 단일 패스 치환 | JSON (통계) |
| `preview <file>` | Markdown 미리보기 | Markdown |
| `fix <file>` | NS 프리픽스 수정 | JSON |

### 단일 패스 최적화

`fill`은 ZIP을 **1회만** 열고 다음을 한번에 처리:
1. 쪼개진 run 병합 (동일 서식 인접 run의 텍스트 합침)
2. 단일 치환 (모든 XML 파일)
3. 순차 치환 (section 파일만)
4. 네임스페이스 교정

### 쪼개진 텍스트 처리

한글 에디터가 동일 서식의 텍스트를 여러 `<run>` 태그로 분리하는 경우 자동 대응:
- **1차**: 동일 `charPrIDRef`의 인접 run 병합 후 치환
- **2차**: run 경계를 넘는 cross-run 패턴 매칭으로 치환

---

## 핵심 규칙

1. **양식 우선**: 사용자 업로드 > assets/ 기본 양식 > HwpxDocument.new()
2. **ZIP-level 치환만 사용**: HwpxDocument.open()은 복잡한 양식 파싱 실패 가능
3. **hwpx_tool.py fill 사용**: 네임스페이스 수정이 자동 포함됨
4. **양식 분석 필수**: 치환 전에 반드시 `analyze`로 텍스트 전수 조사
5. **순차 치환 주의**: 동일 텍스트가 N번이면 N개의 다른 값 필요
6. **날짜 형식**: `2026. 3. 23.` (월/일 앞 0 생략, 온점 사이 공백)
7. **HwpxDocument.new() 후 fix 필수**: `hwpx_tool.py fix <file>` 실행
8. **HWPX 전용**: 레거시 `.hwp`는 처리 불가
9. **글꼴 미포함**: 열람 환경에 해당 글꼴 필요
10. **확인 후 실행**: Phase 3에서 사용자 승인 없이 치환하지 않음

---

## Quick Reference

| 작업 | 방법 |
|------|------|
| 보고서/공문/양식 문서 | 양식 + `hwpx_tool.py fill` |
| 단순 문서 | `HwpxDocument.new()` → `.save()` → `hwpx_tool.py fix` |
| 텍스트 검색/분석 | `hwpx_tool.py analyze` |
| 결과 미리보기 | `hwpx_tool.py preview` |
| 기존 문서 수정 | `hwpx_tool.py fill` (원본 → 수정본) |
