---
allowed-tools: [Read, Write, Edit, Bash, Glob, Grep]
description: "HWPX 한글 양식에 인터랙티브하게 내용을 채워넣는 도구. 'hwpx 채워줘', '한글 양식 채우기', '보고서 양식 작성', 'hwpx fill' 등의 키워드에 사용."
---

# /hwpx-fill — HWPX 양식 인터랙티브 채우기

$ARGUMENTS

이 명령은 `skill/hwpx-fill/` Agent Skill의 Claude Code 진입점이다.

## 도구 경로

스킬 디렉토리: `skill/hwpx-fill/`
통합 스크립트: `skill/hwpx-fill/scripts/hwpx_tool.py`
기본 양식: `skill/hwpx-fill/assets/report-template.hwpx`

## 선행 조건

```bash
pip3 install python-hwpx --break-system-packages
```

## 워크플로우

`skill/hwpx-fill/SKILL.md`의 4단계 워크플로우를 따른다:

1. **Phase 1 (분석)**: `python3 skill/hwpx-fill/scripts/hwpx_tool.py analyze <template.hwpx>`
2. **Phase 2 (값 수집)**: 분석 결과를 표로 보여주고 사용자에게 값 수집
3. **Phase 3 (확인)**: 치환 매핑 전체를 보여주고 사용자 승인
4. **Phase 4 (실행)**: `python3 skill/hwpx-fill/scripts/hwpx_tool.py fill <src> <dst> <json>`

실행 후 미리보기:
```bash
python3 skill/hwpx-fill/scripts/hwpx_tool.py preview <output.hwpx>
```

## 출력 경로

- 사용자 지정 → 해당 경로
- 미지정 → `output/<원본이름>_filled.hwpx`

## 핵심 규칙

- 원본 파일 수정 금지 (fill이 자동으로 src → dst 복사)
- 날짜: `2026. 3. 23.` (월/일 앞 0 생략)
- 사용자 확인 없이 치환 실행 금지
- 상세 규칙은 `skill/hwpx-fill/SKILL.md` 참조
