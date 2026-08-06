# 콘텐츠 JSON 스키마

`scripts/build_report.py`에 `--input`으로 넘기는 JSON의 구조. 리서치한 내용을 이 형태로
정리한 뒤 스크립트를 실행하면 표지 - 목차 - 본문 - 참고자료 순서의 docx가 만들어진다.

## 최상위 필드

| 필드 | 필수 | 설명 |
|---|---|---|
| `title` | O | 보고서 제목 |
| `subtitle` | X | 부제. 있으면 제목 아래 "- 부제 -" 형태로 표시 |
| `author` | X | 작성자. 기본 관례는 `"춘식이 (AI 조사 지원)"` |
| `date` | X | 작성일 (예: `"2026년 8월 6일"`) |
| `doc_type` | X | 문서 구분 (예: `"내부 보고서 (교내 제출용)"`) |
| `main_color` | X | 메인 컬러 hex. CLI의 `--color`가 있으면 그쪽이 우선 |
| `sections` | O | 본문 섹션 배열 (아래 참고) |
| `references` | X | 참고 자료 배열 (아래 참고) |

## `sections` 배열

각 섹션은 다음 필드를 가진다.

| 필드 | 필수 | 설명 |
|---|---|---|
| `title` | O | 섹션 제목. **번호를 직접 포함**시킨다 (예: `"1. 조사 개요"`). 스크립트는 번호를 자동으로 매기지 않는다 — 목차와 본문 제목이 항상 같은 문자열이어야 어긋나지 않기 때문 |
| `blocks` | X | 아래 "블록 타입" 참고 |
| `subsections` | X | 같은 구조(`title` + `blocks`)를 가진 하위 섹션 배열 (예: `"4.1 패션 아이템으로의 진화"`) |

### 블록 타입 (`blocks` 배열의 각 항목)

- `{"type": "paragraph", "text": "..."}` — 일반 본문 단락
- `{"type": "bullets", "items": ["...", "..."]}` — 글머리 기호 목록
- `{"type": "table", "headers": ["...", "..."], "rows": [["...", "..."], ...]}` — 표 (헤더 행은 메인 컬러 배경 + 흰 글씨로 자동 스타일링됨)

## `references` 배열

- `{"text": "출처명, \"기사 제목\"", "url": "https://..."}`
- `url`이 있으면 클릭 가능한 하이퍼링크로 별도 줄에 표시된다. `url`이 없으면 텍스트만 표시.

## 예시 (에코백 시장조사 보고서 발췌)

```json
{
  "title": "Eco-Bag Market Research Report",
  "subtitle": "2026년 최신 트렌드를 중심으로",
  "author": "춘식이 (AI 조사 지원)",
  "date": "2026년 8월 3일",
  "doc_type": "내부 보고서 (교내 제출용)",
  "main_color": "2E74B5",
  "sections": [
    {
      "title": "1. 조사 개요",
      "blocks": [
        {"type": "paragraph", "text": "본 보고서는 2026년을 기준으로 에코백 시장의 최신 동향을 파악하고..."}
      ]
    },
    {
      "title": "4. 2026년 주요 트렌드",
      "blocks": [],
      "subsections": [
        {
          "title": "4.1 패션 아이템으로의 진화",
          "blocks": [
            {"type": "paragraph", "text": "에코백은 단순히 물건을 담는 실용적 도구를 넘어..."}
          ]
        },
        {
          "title": "4.3 소재 혁신",
          "blocks": [
            {"type": "bullets", "items": [
              "재활용 나일론·폴리에스터 소재의 시장 점유율은 약 42% 수준으로 확대",
              "재활용 플라스틱(RPET) 선호도 약 46%"
            ]}
          ]
        }
      ]
    }
  ],
  "references": [
    {"text": "다나와, \"트렌드 CHECK! 2026 에코백 종결판\"", "url": "https://plan.danawa.com/info/?nPlanSeq=12389"}
  ]
}
```
