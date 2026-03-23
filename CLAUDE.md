# HWPX-Fill: 한글 양식 인터랙티브 채우기 도구

## 프로젝트 개요

HWPX 양식 파일(.hwpx)의 원본 구조·서식·레이아웃을 100% 보존하면서 텍스트만 인터랙티브하게 채워넣는 도구.

## 핵심 원리

HWPX 파일은 내부가 ZIP + XML 구조. `HwpxDocument.open()`은 복잡한 양식에서 파싱 실패할 수 있으므로, ZIP 레벨에서 XML 바이트를 직접 문자열 치환하는 전략을 사용한다. 덕분에 원본의 표, 결재란, 색상, 폰트, 레이아웃이 전부 보존된다.

## 프로젝트 구조

```
scripts/
  analyze_template.py    # HWPX 양식 분석 (텍스트 추출 + 분류)
  fill_template.py       # 양식 채우기 (치환 + 네임스페이스 후처리)
templates/               # 사용자 양식 저장소
output/                  # 생성된 문서 출력
ref/                     # 레퍼런스 스킬 (gonggong_hwpxskills)
.claude/commands/
  hwpx-fill.md           # /hwpx-fill 슬래시 커맨드
```

## 사용법

`/hwpx-fill <경로.hwpx>` 로 시작하거나, HWPX 양식 파일 관련 요청을 하면 인터랙티브 모드 진입.

## 기술 스택

- Python 3 + python-hwpx 라이브러리
- ZIP-level XML 텍스트 치환 (HwpxDocument.open() 미사용)
- 네임스페이스 후처리 (한컴 뷰어 호환성)

## 규칙

- 원본 HWPX 파일은 절대 직접 수정하지 않음 — 항상 복사본에서 작업
- 치환 후 반드시 fix_namespaces 실행 (fill_template.py에 내장)
- 날짜 형식: `2026. 3. 23.` (한글 공문서 표준, 월·일 앞 0 생략)
- 보고서 양식 참조: `ref/gonggong_hwpxskills/references/report-style.md`
- 공문서 양식 참조: `ref/gonggong_hwpxskills/references/official-doc-style.md`
