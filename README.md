# VN Office 인사·SCM 출장 관리 시스템

오마이호텔(Oh My Hotel) 베트남 법인의 **인사 통합 관리** 및 **SCM 출장 성과 트래킹**을 위한 단독형 웹 애플리케이션.

**🌐 라이브 사이트**: https://sangis389.github.io/HR-Business-Trip/  
**📦 리포지토리**: https://github.com/sangis389/HR-Business-Trip  
**🚀 현재 버전**: v47

---

## 목차

- [개요](#개요)
- [주요 기능](#주요-기능)
- [화면 구성](#화면-구성)
- [데이터 모델](#데이터-모델)
- [Leave Type 분류](#leave-type-분류)
- [기술 스택](#기술-스택)
- [파일 구조](#파일-구조)
- [로컬 실행](#로컬-실행)
- [데이터 저장 · 백업](#데이터-저장--백업)
- [엑셀 임포트](#엑셀-임포트)
- [배포 워크플로우](#배포-워크플로우)
- [주요 릴리즈 히스토리](#주요-릴리즈-히스토리)

---

## 개요

**대상 조직**: VN Office 총 **46명** (7개 부서)
- Office/SCM (10명) — 출장 담당
- Office/IT (11명), Office/OP (8명), Office/Content (6명), Office/PD (4명), Office/Support (4명)
- KR Manager (3명) — Sang Yoo(SCM Head), Aiden, Justin

**핵심 목표**:
1. **인사 통합 관리** — 근태 · 연차 · 지각 리스크를 사전 감지
2. **SCM 출장 3R 자산화** — 계획 → 실행 → Follow-up 결과 트래킹
3. **법인장 관점 의사결정 지원** — Critical 6 알림 · 부서별 트렌드 · 예측 위젯

---

## 주요 기능

### 🏢 인사(HR) 관리
- **46명 인원 마스터** — 부서 · 직책 · 연차 부여/사용/잔여 · 3직급 관리(Director/Leader/Staff)
- **근태 3-뷰 토글** — 일별 상세 / 인원×월 요약 / 부서×월 요약 (drill-down 지원)
- **11개 Leave Type 분류** — AL / AL/2 / UP / UP/2 / H / CL / MR / FL / BT / SL / MN
- **연차 소진 예측** — 현시점 사용률 기반 연말 시뮬레이션 (미소진/초과 위험 감지)
- **부서별 지각률 트렌드** — 월별 변화 추이 · 3개월 연속 상승 시 ⚠️ 경고
- **인원 공백 캘린더** — 일별 결근+연차 인원수 색상 표시 (5명↑ 빨강)
- **결근/지각 사유 관리** — 사유 입력 · 부서장 리뷰 코멘트 · ✅ 확인 완료 워크플로

### ✈️ SCM 출장 관리
- **5-컬럼 칸반** — DRAFT / REQUESTED / APPROVED / IN_PROGRESS / COMPLETED
- **Analytics 자동 뷰** — 전체 필터 시 KPI 5개 + 지역 분포 + 담당자별 호텔 수 + 담당자×월 매트릭스
- **트립 상세 Recap** — 방문 호텔별 Contract · Purpose · Follow-up · 1주/1개월 후 결과 (5색 뱃지)
- **결과보고 자동 임포트** — Report 시트가 있는 엑셀 드롭 시 hotels[] 자동 파싱·병합
- **트립 제목 통일** — `{City} Biz Trip (DD-DD Mon)` 자동 생성

### 📅 통합 캘린더
- **담당자별 월별 그리드** — 근무(정상/지각/결근) · 출장(BT) · 연차(AL/SL/UP) 통합 표시
- **부서별 필터** — 전체 인원 또는 특정 부서 선택
- **월 셀 hover** — 상세 tooltip (날짜·상태·지각 분·목적지)

### 📊 리포트
- 월별 지각/결근 · 부서별 근태 요약 · 근태 상태 분포 · 트립 상태 분포

---

## 화면 구성

### 사이드바 메뉴
| 아이콘 | 메뉴 | 주요 기능 |
|---|---|---|
| 📊 | **대시보드** | Critical 6 · HR 위젯 · KPI · 부서별 지각 Top 5 · SCM 성과 |
| 👥 | **인원 (VN)** | 46명 리스트 · 부서 필터 · 부여/사용/잔여 컬럼 |
| 🕘 | **근태·연차** | 3탭 (일별 근태 · 휴가 이력 · 담당자별 요약) |
| ✈️ | **SCM 출장** | Analytics + 칸반 · 결과보고 드롭존 |
| 📅 | **통합 캘린더** | 담당자별 월별 근무/출장/연차 그리드 |
| 📈 | **리포트** | 부서별 · 월별 · 상태별 통계 |

---

## 데이터 모델

### Employee (46명)
```javascript
{
  person_id: "00000003",
  name: "Le Thi Kim Anh (Aerum)",
  department: "Office/SCM",
  position: "Staff",               // Director / Leader / Staff / SCM Head
  is_scm: true,
  is_scm_traveler: true,
  annual_leave: 12.0,
  used_leave: 8.0,                 // 실시간 계산 (AL + AL/2 × 0.5)
  remaining_leave: 4.0,            // annual - used
  leave_baseline_date: "2026-06-30"
}
```

### Attendance (4,495건)
```javascript
{
  id: 1,
  person_id: "00000003",
  name, department, date: "2026-01-02",
  check_in: "", check_out: "",
  late_minutes: 0,
  status: "NORMAL" | "LATE" | "ABSENT",
  reason: "",                       // v45 신설
  reason_verified: false,
  reviewer_comment: "",
  note: ""
}
```

### Leave (196건 + auto BT)
```javascript
{
  id: "al-{pid}-{date}",
  person_id, name, date: "2026-07-13",
  type: "AL",                       // 11가지 유형
  days: 1.0,                        // 0.5 for AL/2
  status: "APPROVED",               // DRAFT / REQUESTED / APPROVED / REJECTED
  note: "수기 입력",
  auto: false                       // BT 는 트립에서 자동 생성
}
```

### Trip (34건)
```javascript
{
  id: 1,
  title: "Ha Long Biz Trip (14-16 Jul)",   // 자동 생성 형식 통일
  employee: "Trinh Ngoc Duy (Andy)",
  destination: "Ha Long / Hanoi, VN",
  start_date, end_date,
  purpose: "SCM Business Trip",
  status: "APPROVED",
  cost_planned, cost_actual, currency: "VND",
  hotels: [                                 // 방문 호텔 Recap
    {
      hotel: "Wyndham Legend Halong",
      contract: "Shared / HTL",
      contact: "Mr. Truyen - DOS",
      purpose, summary, followup,
      result_1w, result_1m,
      status: "DONE" | "REQUESTED" | "IN_PROGRESS" | "REJECTED" | "PENDING"
    }
  ],
  expense_vnd, source_file, outcome, roi, notes
}
```

---

## Leave Type 분류

| 코드 | 명칭 | 연차 차감 | 유급 | 그룹 | 색상 |
|---|---|---|---|---|---|
| **AL** | Annual Leave | ✅ 1.0 | 유급 | annual | 🔵 파랑 |
| **AL/2** | Half Annual Leave | ✅ 0.5 | 유급 | annual | 🔵 하늘 |
| **UP** | Unpaid Leave | ✗ | 무급 | unpaid | ⚪ 회색 |
| **UP/2** | Half Unpaid Leave | ✗ | 무급 | unpaid | ⚪ 연회색 |
| **H** | Holiday | ✗ | 유급 | paid_pl | 🟢 초록 |
| **CL** | Compensation Leave | ✗ | 유급 | paid_pl | 🟢 밝은초록 |
| **MR** | Marriage Leave | ✗ | 유급 | paid_pl | 🩷 핑크 |
| **FL** | Funeral Leave | ✗ | 유급 | paid_pl | ⚫ 다크그레이 |
| **BT** | Business Trip | ✗ | 유급 | paid_pl | 🟠 주황 (트립에서 자동 생성) |
| **SL** | Sick Leave | ✗ | 무급 | sick | 🔴 빨강 |
| **MN** | Maternity Leave | ✗ | 무급 (SI) | maternity | 🌸 라이트 핑크 |

**잔여 계산 공식**: `remaining = annual_leave - Σ(AL days) - Σ(AL/2 days × 0.5)`

---

## 기술 스택

- **Frontend**: Vanilla JavaScript (SPA 방식, 단일 render loop)
- **차트**: 순수 CSS + SVG (라이브러리 없음)
- **엑셀 파싱**: [SheetJS](https://sheetjs.com/) v0.18.5 (CDN)
- **저장소**: 브라우저 localStorage + `data.json` 시드
- **배포**: GitHub Pages via GitHub Actions
- **의존성**: 없음 (완전 static)

---

## 파일 구조

```
SCM 인사 출장 관리/
├── index.html                    # 앱 진입점
├── app.js                        # 애플리케이션 로직 (~2,700줄)
├── styles.css                    # 스타일
├── data.json                     # 시드 데이터 (employees + attendance + leaves + trips)
├── README.md                     # 이 파일
├── LICENSE                       # MIT
├── .gitignore
├── .github/workflows/
│   └── pages.yml                 # GitHub Pages 자동 배포
├── docs/
│   └── specs/scm-integrated-dashboard/
│       └── ko/                   # FR/NFR 스펙 문서
├── scripts/
│   └── deploy-config.txt         # PAT (gitignore)
└── data-source/                  # 원본 엑셀 (근태 · 출장 계획서/결과보고)
```

---

## 로컬 실행

```bash
# Python
python -m http.server 8000

# Node
npx serve
```

접속: http://localhost:8000

> `fetch("data.json")` 이 `file://` 프로토콜에서 CORS 오류가 날 수 있어 로컬 서버 권장.

---

## 데이터 저장 · 백업

- **localStorage**: `vn-office-v47` 키로 저장 (버전 업데이트 시 자동 무효화)
- **캐시 초기화**: 우측 상단 `데이터 초기화` 버튼 → 서버의 `data.json` 리로드
- **JSON 백업**: 사이드바 하단 `전체 백업 (JSON)` 클릭 → 파일 다운로드 (employees + attendance + leaves + trips 전체)

---

## 엑셀 임포트

### 근태 임포트
근태·연차 페이지의 드래그 존에 **KEYWATCH 형식 xlsx** 드롭
- 자동 파싱: Person ID · Date · Check-in · Check-out
- 08:15 규칙으로 지각 자동 계산

### 출장 임포트 (계획서 · 결과보고 자동 인식)
SCM 출장 페이지 드래그 존에 xlsx 드롭 → **Report 시트 유무로 자동 분기**:

**계획서 모드** (Report 시트 없음 또는 hotels < 3):
- Plan/Schedule 시트에서 날짜 · 호텔 리스트 파싱
- Expense 시트에서 경비 추출
- status = DRAFT 또는 APPROVED

**결과보고 모드** (Report 시트에 hotels ≥ 3):
- 파일명에서 담당자 · 기간 · 목적지 자동 파싱
- 기존 트립 매칭 (담당자 + 시작일 ±2일) → 병합 또는 신규 생성
- Report 시트에서 hotels[] 자동 파싱 (Contract · Purpose · Follow-up · 결과)
- 결과 텍스트 → 5색 상태 뱃지 자동 분류

---

## 배포 워크플로우

1. 로컬에서 `app.js` / `data.json` / `styles.css` 수정
2. `STORAGE_KEY` 버전 bump (v47 → v48) — 브라우저 캐시 자동 무효화
3. GitHub 로 push (main 브랜치)
4. GitHub Actions (`.github/workflows/pages.yml`) 자동 실행 → GitHub Pages 배포
5. 1-3분 후 라이브 사이트 반영

**보안**: PAT 은 `scripts/deploy-config.txt` (gitignore) 에 로컬 저장. Git URL 에 임시 주입 후 원격 리셋.

---

## 주요 릴리즈 히스토리

| 버전 | 주요 변경 사항 |
|---|---|
| **v47** | Aiden 6/22-24 트립 추가 |
| v46 | Sang Yoo 6/23-30 업데이트 · Aiden 6/29-7/2 신규 트립 |
| **v45** | **HR 관리 강화** — 정합성 audit 12건 + 사유 입력 + 연차 예측/지각 트렌드/공백 캘린더 위젯 |
| v44 | Mike Da Nang 계획서 추가 (7/28-30, 21곳) |
| v43 | Monti · MiA 7월 AL 반영 |
| v42 | Nhi 7/13 AL + 트립 기간 중복 AL 정리 |
| v41 | 트립 제목 통일 형식 `{City} Biz Trip (DD-DD Mon)` |
| v40 | DK Hanoi 계획서 (7/14-16, 17곳) |
| **v39** | **근태 3-뷰 토글** (일별/인원×월/부서×월) + drill-down |
| v38 | SCM 뱃지 제거 (Head 뱃지만 유지) |
| v36-37 | TBD → Abroad 전환 및 개별 정정 |
| v35 | 직책 업데이트 (Director 4, Leader 5) |
| **v34** | 월/유형 필터 드롭다운 전환 |
| v32-33 | leaves 필드 load/save 버그 수정 + 캘린더 중복 제거 |
| **v27** | **통합 캘린더 뷰** 신설 |
| **v25** | **근태·연차 통합 (3탭)** + 11가지 Leave Type |
| **v20** | **SCM 출장 Analytics 뷰** (KPI + 지역/담당자 차트 + 매트릭스) |
| v19 | Andy 출장 결과보고 3건 반영 (Hanoi) |
| v18 | Lukas(Do Xuan Huy Hoang) 신규 SCM · Thu/Slena 출장 제외 |
| **v17** | **트립 Recap** — 방문 호텔 리스트 · 성과 KPI |
| v15-16 | 대시보드 SCM 출장 요약 카드 · 담당자 필터 |
| v14 | SCM 출장 월별 필터 |
| v13 | Timesheet BT 파싱 · 자동 트립 생성 |
| v10-12 | 연차 배정 · 계산식 정정 (baseline + 미래 예측) |
| v8-9 | Office/PD 부서 신설 · Sang Yoo SCM Head |
| v5-7 | 지각 기준 08:15 · 데이터 재계산 |
| v1-4 | 초기 데이터 로드 · 5개 뷰 구조 확립 |

---

## 스펙 문서 (Round 2 최종)

**Planner 점수**: 9/10 · **Tester 점수**: 8/10 · **상태**: FINALIZED (2026-07-06)

주요 확정 사항:
- FR-503 영업일 기준 D+1/D+3/D+7 에스컬레이션
- FR-202/203 반차+지각 status 규칙
- FR-704(b) 화면 vs 파일 마스킹 정책
- FR-405 amendment 낙관 잠금 상호작용
- FR-101 IANA TZ 문자열 명시
- FR-503 3단계 색상 코딩 (파랑/노랑/빨강)

Sprint 유예 항목 6건 (dangling manager_id, Open Questions 재검토, TC-900 경계 세분화 등)

상세: `docs/specs/scm-integrated-dashboard/ko/scm-integrated-dashboard-spec.md`

---

## 라이선스

MIT

---

**Contact**: Global_SCM@ohmyhotel.com
