# 📋 기획서 분석 에이전트 (Planning-Analyst)

PPT/PDF 기획서를 QA 관점으로 분석해 **명세 누락·모호성·이미지 불일치**를 질의서로 산출하는 에이전트

---

## 🚀 실행 방법

### 기본 명령
```bash
/analyze-plan
```
- `docs/` 폴더에서 `.pptx` 또는 `.pdf` 파일을 자동으로 찾아 분석
- 파일이 정확히 1개만 있으면 자동 선택

### 특정 파일 지정
```bash
/analyze-plan docs/기획서.pptx
```

---

## 📊 처리 흐름 (4+1 단계)

### 0️⃣ 대상 파일 결정
- `docs/` 폴더에서 `.pptx` 또는 `.pdf` 자동 검색
- 1개 발견 시 자동 선택, 2개 이상이면 사용자에게 선택 요청

### 1️⃣ 산출물 정리
**삭제 대상:**
- `output/chunks/`
- `output/manifest.json`
- `output/qa-queries.json`
- `output/qa-query-sheet.md`
- `output/qa-query-sheet.csv`

**유지 대상:**
- `output/partial/` — 재실행 시 중단점부터 이어서 분석하기 위해 기본 유지
  - "처음부터 다시" 또는 "새로 분석" 명시 시만 삭제

### 2️⃣ 기획서 분석 (Planning-Analyst 에이전트)

#### **역할**
10년차 QA Engineer 페르소나로 기획서를 요약하지 않고, **쓰여 있지 않은 것**, **모호한 것**, **충돌하는 것**을 찾아 기획자가 회신할 수 있는 질의로 만든다.

#### **내부 절차**

**1. 청크 준비**
```bash
python scripts/extract_doc.py <파일경로> -o output --chunk-size 8 --overlap 1
```
- PPT/PDF를 8개 항목씩 청크로 분할
- `output/chunks/` 에 C001.md, C002.md, ... 생성
- `manifest.json` 생성 (needs_vision 플래그 포함)

**2. 청크 분석 (요소 유형별 매핑)**

각 청크를 읽고, **해당하는 요소만** 검토:

| 요소 유형 | 검토 카테고리 |
|---------|-----------|
| 버튼/링크 | A-13 상호작용, A-04 상태조합, A-05 중복실행 |
| 입력창 | A-01 경계값, A-02 빈값, A-03 예외문자 |
| 필터/드롭다운 | A-06 조건결합, A-09 상태유지 |
| 목록/카드 | A-01 0건처리, A-07 정렬 |
| 텍스트 표시 | A-11 문구표기, A-12 용어정의 |
| 화면 전체 | A-08 에러처리 (1회만) |

**중요:**
- 각 카테고리의 구체적 점검 질문은 `references/checklist-A-spec.md` 참조
- **A-13(상호작용)을 버튼에 가장 먼저 적용** — "클릭 시 ~한다"만 있고 결과 표시 방식(모달/인라인/새페이지/자동재생) 없으면 질의 생성
- **판정 결과('준수', '해당없음')를 기록하지 않는다** — 실제 질의만 기록

**3. 이미지 대조** (needs_vision=true인 경우)
- 슬라이드당 1회만 이미지를 열어 텍스트 설명과 대조
- 불일치 발견 시 별도 질의 생성:
  - 설명에는 있는데 목업에 없음
  - 버튼 문구가 다름
  - 콜아웃 번호 누락

**4. 청크 통합** (청크 2개 이상일 때만)
- 중복 질의 병합 (출처 모두 나열, 반복 발견 시 우선순위 한 단계 상향)
- 문서 간 충돌 탐지:
  - 같은 항목의 다른 수치/정책 (예: 세션 타임아웃 30분 vs 60분)
  - 같은 대상의 다른 용어 (예: 삭제 vs 제거 vs 폐기)
- 커버리지 점검
- 통합 채번 + 우선순위 정렬
- ※ 청크가 1개면 이 단계 생략

**부분 저장** (청크 3개 이상)
- 각 청크 분석 완료 시 `output/partial/{chunk_id}.json` 즉시 저장
- 다음 청크로 넘어가기 전에 반드시 완료
- 재실행 시: 이미 저장된 청크 건너뛰고 없는 청크부터 이어서 분석

**5. 출력: qa-queries.json**

```json
{
  "document": "파일명",
  "generated_at": "YYYY-MM-DD HH:MM",
  "queries": [{
    "id": "Q-001",
    "element_no": "④",
    "category": "입력 검증",
    "case": "검색창 최대 입력 길이 미정",
    "quote": "[슬라이드 1] ※ 최대 입력 길이 정책 없음",
    "risk": "입력 검증 기준이 없어 TC 작성 불가, XSS 취약점 가능",
    "expected_action": "최대 50자로 제한하는 것을 권고합니다",
    "priority": "필수|높음|중간",
    "owner": "기획|개발|디자인|보안",
    "status": "대기"
  }],
  "conflicts": [],
  "accessibility": [],
  "unreadable": []
}
```

**필드 규칙:**
- `element_no`: 기획서의 번호 콜아웃(①②③...). 없으면 빈 문자열
- `category`: 자연어로 쓴다. A-01 같은 내부 코드는 쓰지 않음
- `quote`: 기획서 원문 인용. **인용할 수 없으면 그 질의를 만들지 않는다**
- `expected_action`: 구체적 대안 제시. "확정 필요", "기획자 확인" 같은 문구 금지

**6. 검증 및 렌더링**

```bash
python scripts/validate_queries.py output/qa-queries.json
python scripts/render_query_sheet.py output/qa-queries.json output
```
- validate_queries.py: 오류 검증 (누락된 필드 등)
- render_query_sheet.py: qa-queries.json → md/csv 생성

### 3️⃣ 최종 보고
- 분석 대상 파일명
- 처리된 청크 수
- 질의 개수 (우선순위별: 필수/높음/중간)
- 문서 간 충돌 개수
- 접근성(KWCAG) 확인 결과
- output/ 에 생성된 최종 파일 목록과 각 파일 크기

### 4️⃣ md/csv 생성
```bash
python scripts/render_query_sheet.py output/qa-queries.json output
```
- `qa-query-sheet.md` — 마크다운 형식 질의서
- `qa-query-sheet.csv` — Excel 호환 CSV (UTF-8 BOM)

---

## 📁 산출물

| 파일 | 설명 |
|-----|------|
| `qa-query-sheet.md` | 마크다운 형식 질의서 (1. 확인사항 / 2. 충돌 / 3. 접근성 / 4. 분석불가) |
| `qa-query-sheet.csv` | Excel 호환 CSV (UTF-8 BOM) |
| `qa-queries.json` | 분석 데이터 소스 |
| `chunks/C001.md~` | 청크 파일들 (extract_doc.py 생성) |
| `manifest.json` | 청크 메타데이터 |
| `partial/*.json` | 청크별 분석 (재실행 대비, 유지) |

---

## ⚡ 성능 비교

| 분석 방식 | 시간 | 질의 수 | 특징 |
|---------|------|--------|------|
| **텍스트만** | 3.7분 | 19건 | 빠름, 이미지 검토 제외 |
| **이미지 포함** | 4.5분 | 19건 | 이미지 대조 수행 |
| **이전 버전** | 15분 | 23건 | 구 스키마, 판정 결과 기록 |

**개선 효과:**
- 요소 유형별 매핑 → 불필요한 검토 제외 → 효율성 ↑
- 판정 결과 기록 제거 → 파일 크기 50% 감소
- 부분 저장 → 중단 후 재실행 가능

---

## 🎯 주요 특징

✅ **요소 유형별 매핑**
- 불필요한 카테고리 자동 제외
- 검토 범위 명확화

✅ **판정 결과 기록 금지**
- '준수', '해당없음' 같은 판정 제거
- 실제 질의만 기록

✅ **부분 저장 & 재실행**
- 청크 3개 이상 시 중단점 저장
- 중단 후 재실행해도 처음부터 하지 않음

✅ **이미지 판독**
- 슬라이드 1회 열어서 텍스트-이미지 대조
- 불일치 자동 감지

✅ **조건부 접근성 검토**
- 기본: 접근성 검토 생략
- "접근성도 봐줘" 명시 시만 수행

---

## 🚫 금지 사항

❌ 테스트 케이스나 테스트 전략 작성
- tc-writer / test-planner 역할

❌ 기획서에 없는 동작을 요구사항으로 창작
- 질의로 배출

❌ 판정 기록("해당없음", "준수", "명시 없음")
- 발견만 기록

❌ 중간 분석 파일 생성
- output/analysis/*.json 만들지 않음

❌ 총평("대체로 잘 작성됨")
- 구체적 판정과 근거만

---

## 💡 팁

### 빠른 분석이 필요한 경우
```bash
/analyze-plan 파일명 이미지 판독은 건너뛰고 텍스트만으로 분석해줘
```
→ 약 3.7분 (텍스트만)

### 정밀한 분석이 필요한 경우
```bash
/analyze-plan 파일명
```
→ 약 4.5분 (이미지 판독 포함)

### 대용량 기획서 (청크 3개 이상) 분석
1. 첫 실행: 전체 청크 분석 시작
2. 중단 시: partial/ 폴더에 완료된 청크 저장
3. 재실행: 미완료 청크부터 이어서 분석

---

## 🔧 설정 파일

### `.claude/agents/planning-analyst.md`
- 에이전트의 역할, 절차, 규칙 정의
- 요소 유형별 매핑 규칙
- 판정 결과 기록 금지 규칙

### `.claude/commands/analyze-plan.md`
- 4+1단계 파이프라인 정의
- output/partial/ 조건부 삭제 규칙

### `scripts/extract_doc.py`
- PPTX/PDF 청크 추출
- manifest.json 생성

### `scripts/validate_queries.py`
- qa-queries.json 검증

### `scripts/render_query_sheet.py`
- qa-queries.json → md/csv 렌더링

### `.claude/settings.local.json`
- 권한 설정 (Bash, PowerShell)

---

## 📝 예시

### 분석 요청
```bash
/analyze-plan docs/OTT_콘텐츠찾기_화면설계서.pptx
```

### 출력 예시
```
✅ 0단계: docs/OTT_콘텐츠찾기_화면설계서.pptx 선택
✅ 1단계: 산출물 정리 완료
⏳ 2단계 진행: planning-analyst 에이전트 실행 (약 4.5분)

[완료 후]

📊 최종 보고
- 분석 대상 파일: OTT_콘텐츠찾기_화면설계서.pptx
- 처리된 청크 수: 1개
- 질의 개수: 19건 (필수 7 / 높음 6 / 중간 6)
- 문서 간 충돌: 0건
- 접근성 확인: 미시행

📁 산출 파일
- qa-query-sheet.md: 10.4KB
- qa-query-sheet.csv: 10.0KB
- qa-queries.json: 13.6KB
```

---

**작성일:** 2026-08-06  
**에이전트 버전:** planning-analyst (슬림 버전)  
**최적화:** 요소 유형별 매핑, 판정 결과 제거, 부분 저장
