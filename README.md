# TUB Voting Dashboard + PPTX Alignment

Azure Technical Update Briefing(TUB) 월간 PPTX에서 **Voting Dashboard(Excel)** 결과를 읽어 선정된 슬라이드만 추출하고, 담당자(Area)별로 PowerPoint Section을 자동 구성하는 도구 모음.

## 워크플로우

```
Voting Dashboard (Excel)          PPTX (원본)
       │                              │
       └──── tub_agent.py ────────────┘
                   │
                   ▼
         _sectioned.pptx
   (Opening → IH → CH → YJ → 나머지 → Closeout)
```

1. **Voting Dashboard (Excel)** 에서 투표 결과 확인 → 음영 처리된 항목 = 선정
2. **PPTX 필터링** — 비선정 슬라이드 삭제
3. **Section 구성** — Area(IH / CH / YJ / 나머지)별 재정렬 + PowerPoint Section 생성

## 설치

```bash
pip install -r requirements.txt
```

## 스크립트

| 파일 | 설명 |
|---|---|
| `tub_agent.py` | **메인** — Excel Voting Dashboard 연동, 슬라이드 필터·정렬·섹션 자동 구성 |
| `tub_build_sections.py` | 하드코딩된 키워드 기반, 선정 슬라이드 삭제 + Area별 Section 재정렬 |
| `hide_slides_2026_04.py` | KEEP 리스트 기반 슬라이드 숨김 처리 (간단 버전) |

## 사용법

### 기본 실행 (Excel 자동 탐색)

PPTX와 같은 폴더에 `*Voting*Dashboard*.xlsx`가 있으면 자동으로 찾습니다.

```bash
python tub_agent.py "2026_04 - Azure-TUB - 복사본.pptx"
```

### Excel / 시트 직접 지정

```bash
python tub_agent.py "2026_04 - Azure-TUB - 복사본.pptx" \
  --excel "Azure_TUB_Voting Dashboard.xlsx" \
  --sheet "2026-04_result"
```

### 출력 경로 지정

```bash
python tub_agent.py input.pptx --output result.pptx
```

### 보조 스크립트

```bash
# 슬라이드 숨김만 (삭제 X)
python hide_slides_2026_04.py "2026_04 - Azure-TUB - 복사본.pptx"

# 하드코딩 키워드 기반 Section 구성
python tub_build_sections.py "2026_04 - Azure-TUB - 복사본.pptx"
```

## 출력

| 파일 | 생성 스크립트 | 내용 |
|---|---|---|
| `*_sectioned.pptx` | `tub_agent.py`, `tub_build_sections.py` | 비선정 삭제 + Area별 Section 정렬 |
| `*_filtered.pptx` | `hide_slides_2026_04.py` | 비선정 슬라이드 숨김 (삭제 아님) |

## 동작 상세

### Excel 파싱 (`tub_agent.py`)
- `_result` 접미사가 있는 시트를 자동 선택 (최신 우선)
- 셀 배경색(음영)이 있는 행 = 선정 항목
- C열에서 담당 Area(IH/CH/YJ) 추출, 없으면 `나머지`

### 슬라이드 분류
- **Opening**: 표지 ~ Agenda (포함)
- **IH / CH / YJ / 나머지**: Excel 선정 항목 + 해당 카테고리 헤더
- **Closeout**: Thank You / Q&A
- **삭제**: 위 어디에도 해당하지 않는 슬라이드

### Section 순서
`Opening → IH → CH → YJ → 나머지 → Closeout`