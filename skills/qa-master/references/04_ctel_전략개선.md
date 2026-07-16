# ISTQB Expert Level(CTEL) 참고 지식 — 테스트 전략·프로세스 개선

> 목적: 조직/프로젝트 차원에서 테스트를 전략적으로 운영·개선하는 규칙을 CTEL 두 모듈(Test Management, Improving the Test Process) 기준으로 정리
> 성격: 보조 지식 문서. 웹 조사 + 표준 지식 기반. 확인 안 된 항목은 문서 끝에 별도 표기
> 작성 기준일: 2026-07-15

---

## 0. Expert Level 현재 제공 상태 (istqb.org 기준)

- Expert Level은 단종되지 않음. 현재 Core 스트림에서 Test Management와 Improving the Test Process 두 갈래로 제공 중
- ISTQB 유일하게 유효기간이 있는 레벨. 인증 유효기간 7년이며 갱신(Extension) 필요
- 응시 전제조건: CTFL(Foundation) + CTAL-TM(Advanced Test Manager) 보유, 실무 테스트 경험 최소 5년·테스트 관리 경험 최소 2년
- 2024년에 "단종"된 것은 Expert Level 자체가 아니라 구버전(예: CTFL 3.1) 시험판임 (영어 2024-05, 기타 언어 2024-11 종료). Expert Level과 혼동 주의

### 모듈 구성 (각 파트가 별도 시험, 전 파트 취득 시 완전 인증)

| 인증 | 파트 | 코드 |
|------|------|------|
| CTEL-TM (Test Management) | Strategic Test Management | CTEL-TM-SM |
| | Operational Test Management | CTEL-TM-OTM |
| | Managing the Test Team | CTEL-TM-MTT |
| CTEL-ITP (Improving the Test Process) | Assessing the Test Process | CTEL-ITP-ATP |
| | Implementing Test Process Improvement | CTEL-ITP-ITPI |

- 각 파트 시험 구성(공개 정보 기준): 객관식 14~16문항 + 에세이(3중 2문항) + 135분

---

## 1. 테스트 정책·전략·접근법의 3계층 구조

조직→제품라인→프로젝트로 내려가며 추상도가 낮아지고 구체성이 높아지는 하향 계층 구조.

| 계층 | 소유 주체/수준 | 담는 내용 | 성격 |
|------|----------------|-----------|------|
| 테스트 정책(Test Policy) | 조직 전체 | 테스트의 목적·미션, 테스트의 가치/투자 관점, 품질에 대한 조직 태도, 프로세스 개선 방침, 테스트 평가 기준 | 왜(Why) 테스트하는가 — 최상위, 짧고 안정적 |
| 테스트 전략(Test Strategy) | 조직·제품라인 | 테스트 레벨/타입, 리스크 기반 접근, 진입·종료 기준, 자동화 방침, 테스트 환경·데이터, 재현/회귀 방침, 도구 표준 | 어떻게(How) 일반적으로 테스트하는가 — 프로젝트 무관 일반 규칙 |
| 테스트 접근법(Test Approach) | 개별 프로젝트 | 전략을 특정 프로젝트에 맞게 재단(tailoring)한 실행안. 테스트 계획서(Master/Level Test Plan)에 반영 | 이번 프로젝트에서 실제로 어떻게 할 것인가 — 구체·가변 |

### 전략 유형(Test Strategy Types) — 접근법 선택 시 조합 사용

- 분석적(Analytical): 리스크 기반, 요구사항 기반 등 분석에 근거
- 모델 기반(Model-based): 상태전이·사용프로파일 등 모델에 근거
- 방법론적(Methodical): 체크리스트·표준 품질특성 등 정형 목록 활용
- 프로세스 준수형(Process/Standard-compliant): IEEE·애자일 등 외부 표준 준수
- 지시형(Directed/Consultative): 이해관계자·전문가 의견에 근거
- 회귀 회피형(Regression-averse): 재사용·자동화로 회귀 방지 중심
- 반응형(Reactive): 사전 설계 최소화, 실제 시스템 대면 후 대응(예: 탐색적 테스트)

> 실무 규칙: 정책은 거의 안 바뀌고, 전략은 조직 표준으로 재사용하며, 접근법만 프로젝트마다 재단한다. 프로젝트 문서(테스트 계획서)는 "전략의 어느 부분을 왜 재단했는지"를 명시해야 추적 가능.

---

## 2. 전략적 테스트 관리 (CTEL-TM-SM)

### 2-1. 테스트 조직 구성 모델

- 조직 내 테스트 기능의 배치 형태를 선택하는 문제
  - 프로젝트 내장형(개발팀에 테스터 포함): 밀착·빠른 피드백, 독립성 약함
  - 독립 테스트팀/센터(TCoE, Test Center of Excellence): 표준화·독립성 강함, 밀착도 낮음
  - 혼합형: 도메인 지식은 내장, 표준/도구/전문 테스트는 중앙 지원
- 판단 기준: 독립성 vs 도메인 밀착, 표준화 수준, 규모의 경제, 커뮤니케이션 비용

### 2-2. 아웃소싱/인소싱 판단 (소싱 전략)

- 유형: 인소싱(사내), 아웃소싱(외부위탁), 오프쇼어/니어쇼어(원격지), 혼합
- 판단 시 고려 요소
  - 비용: 단가뿐 아니라 인수인계·조정·품질검증 비용 총액(TCO)
  - 핵심역량 여부: 제품 지식·보안 민감도가 높으면 사내 유지
  - 통제·품질검증: 산출물 품질을 어떻게 계약·측정·검증할지(SLA, 품질지표)
  - 커뮤니케이션·시차·문화·언어 리스크
  - 지식 이전·의존성 고착(vendor lock-in) 리스크
- 규칙: 아웃소싱은 "책임까지 위탁되지 않는다". 품질 목표·수용 기준·측정 방법을 계약과 리포팅으로 발주 측이 통제

### 2-3. 테스트의 가치(투자 대비 가치) 설명법

- 테스트를 비용이 아니라 리스크 감소·의사결정 지원 활동으로 설명
- 가치 논거의 축
  - 결함의 조기 발견 비용 이점(후행 단계일수록 수정비용 급증)
  - 리스크 완화: 잔존 리스크를 가시화해 릴리스 의사결정 근거 제공
  - 재작업/장애/평판 손실 등 회피 비용
- 경영진 커뮤니케이션 규칙: 기술 지표가 아니라 비용·리스크·일정 언어로 환산해 제시

---

## 3. 운영 테스트 관리 (CTEL-TM-OTM)

### 3-1. 여러 프로젝트/프로그램에 걸친 테스트 관리

- 다수 프로젝트를 동시에 다룰 때: 자원(인력·환경·도구) 공유·경합 조정, 우선순위·리스크 기반 배분
- 조직 표준(정책·전략)을 각 프로젝트 접근법으로 일관되게 전개하고, 프로젝트 간 교훈(lessons learned)을 표준으로 환류
- 프로그램/포트폴리오 수준의 진척·리스크 통합 리포팅

### 3-2. 제3자(3rd party) 관계 관리 — OTM 핵심 주제

Test Manager가 외부 공급자를 다룰 때 필요한 역량 4축:
- 계약(Contractual): 범위·수용기준·품질 SLA·책임 경계 명문화
- 커뮤니케이션(Communication): 정례 보고·에스컬레이션 경로·언어/시차 관리
- 통합(Integration): 외부 산출물을 사내 프로세스·환경·도구체인에 통합
- 품질 검증(Verification of quality): 넘어온 산출물의 품질을 독립 검증(수용 테스트·지표 검토)

### 3-3. 이해관계자 커뮤니케이션

- 이해관계자별로 관심사가 다름 → 메시지를 청중에 맞게 재단
  - 경영진: 리스크·비용·릴리스 가부
  - 개발/PM: 결함 상세·차단 이슈·일정 영향
  - 고객/현업: 수용 기준 충족 여부·업무 영향
- 규칙: 나쁜 소식일수록 조기·사실 기반으로 전달, 지표는 해석(권고)과 함께 제시

---

## 4. 테스트 팀 관리 (CTEL-TM-MTT)

- Test Manager의 사람 관리 역량(팀 구축·육성·리딩)에 초점
- 다루는 영역(일반)
  - 팀 구성·채용: 필요한 스킬 정의, 개인/팀 역량 균형
  - 동기부여·유지: 인정·성장경로·자율성
  - 스킬 개발: 개인별 역량 진단과 교육 계획
  - 팀 다이내믹스: 분산/원격 팀, 문화 다양성, 갈등 관리
  - 리더십: 위임, 성과 피드백, 방향 제시

> 상세 하위 주제는 공식 실러버스 원문 대조 필요(아래 확인 불가 항목 참조).

---

## 5. 테스트 프로세스 개선 (CTEL-ITP)

### 5-1. 개선 접근법 3분류

| 접근법 | 성격 | 대표 모델 |
|--------|------|-----------|
| 프로세스/모델 기반(Process·Model-based) | 성숙도 단계·프로세스영역 기준으로 갭 진단 후 상향 | TMMi, TPI NEXT |
| 내용 기반(Content-based) | 정해진 "좋은 테스트 실천" 참조모델과 대조 | CTP, STEP |
| 분석 기반(Analytical-based) | 데이터로 원인 규명·측정 주도 | 인과분석(Causal Analysis), GQM(Goal-Question-Metric) |

### 5-2. TMMi — 5단계 성숙도 (Test Maturity Model integration)

- CMMI를 테스트 관점으로 보완하는 단계형(staged) 성숙도 모델
- 상위 단계로 가려면 하위 단계 프로세스영역을 충족해야 함(스테이지드 진행)

| 레벨 | 이름 | 특징 | 프로세스 영역(대표) |
|------|------|------|----------------------|
| 1 | Initial(초기) | 테스트=디버깅, 비정형·비문서화. 프로세스 영역 없음 | — |
| 2 | Managed(관리됨) | 테스트를 디버깅과 분리, 계획·설계·실행·통제의 기본 체계 확립 | 테스트 정책·전략 / 테스트 계획 / 테스트 모니터링·통제 / 테스트 설계·실행 / 테스트 환경 |
| 3 | Defined(정의됨) | 테스트를 개발 생명주기에 통합, 조직 표준 프로세스·전담 테스트 조직 | 테스트 조직 / 테스트 교육 프로그램 / 테스트 생명주기·통합 / 비기능 테스트 / 동료 검토(Peer Reviews) |
| 4 | Measured(측정됨) | 테스트를 정량적으로 측정·통제, 제품 품질을 정량 평가 | 테스트 측정 / 제품(소프트웨어) 품질 평가 / 고급 동료 검토(Advanced Reviews) |
| 5 | Optimization(최적화) | 데이터로 결함 예방·프로세스 지속 최적화(통계적 통제) | 결함 예방 / 품질 통제 / 테스트 프로세스 최적화 |

- 진행 규칙(일반적으로 알려진 관행): 각 프로세스영역 목표를 상당 수준 충족해야 다음 레벨 인정(흔히 "약 85% 이상 달성" 관행이 인용됨 — 5-7 참조)

### 5-3. TPI NEXT — 핵심 영역/성숙도

- Sogeti의 TPI를 발전시킨 모델. 특정 개발방법론에 독립적이며, 조직 목표에 맞춰 목표(objectives) 재단 가능
- 16개 핵심 영역(Key Areas)을 각각 성숙시키는 방식(연속형에 가까움)
- 성숙도 4단계: Initial → Controlled → Efficient → Optimizing
- 도구 요소: 성숙도 매트릭스(Maturity Matrix), 체크포인트(Checkpoint), 클러스터(Cluster, 개선 순서를 묶어 제시), 개선 제안(Improvement suggestions), 조력자(Enablers)
- 16개 핵심 영역(일반적으로 정리되는 목록): 이해관계자 관여(Stakeholder commitment), 관여 정도(Degree of involvement), 테스트 전략, 테스트 조직, 커뮤니케이션, 리포팅, 테스트 프로세스 관리, 견적·계획, 지표(Metrics), 결함 관리, 테스트웨어 관리, 방법론 실천, 테스터 전문성, 테스트 케이스 설계, 테스트 도구, 테스트 환경
  - (개별 명칭은 번역/판본에 따라 표기 차 있음 — 5-7 참조)

### 5-4. CTP — Critical Testing Processes (내용 기반)

- Rex Black이 제시한 내용 참조 모델. "정해진 임계 테스트 프로세스들"이 제대로 되면 성공, 못 되면 실패한다는 전제
- 12개 임계 프로세스를 식별하고, 산업 지표(metrics)와 벤치마킹으로 현 수준을 비교
- 특정 성숙도 순서를 강제하지 않아 어떤 SDLC에도 통합 가능(유연·경량)
- 프로세스는 통상 계획(plan)·준비(prepare)·수행(perform)·개선(perfect) 성격으로 묶임
  - (12개 프로세스의 정확한 개별 명칭은 원문 대조 필요 — 5-7 참조)

### 5-5. STEP — Systematic Test and Evaluation Process (내용 기반)

- 요구사항 기반 전략을 강조하는 내용 참조 모델. "코딩 전에 테스트(테스트를 코딩보다 먼저 설계)" 원칙
- 개선에 정해진 순서를 강제하지 않음(내용 기반 특성)
- 특징: 조기 테스트 설계, 테스터·개발자 협업, 초기 단계 결함 식별
- 통상 3대 국면으로 설명: 계획(Plan) → 획득/분석·설계(Acquire) → 측정/실행·평가(Measure)
  - (국면 명칭·세부는 판본 차 있음 — 5-7 참조)

### 5-6. 개선 사이클 — IDEAL

SEI가 정의한 프로세스 개선 사이클. ITP에서 개선 프로젝트 운영 프레임으로 활용.

| 단계 | 뜻 | 핵심 활동 |
|------|-----|-----------|
| I — Initiating(착수) | 개선의 착수 | 개선 배경·목표·후원 확보, 개선 인프라 준비 |
| D — Diagnosing(진단) | 현 상태 파악 | 현행 프로세스 평가(assessment), 갭 분석, 개선 권고 도출 |
| E — Establishing(수립) | 계획 수립 | 우선순위 설정, 개선 접근·계획·측정 지표 정의 |
| A — Acting(실행) | 개선 실행 | 해법 개발·파일럿·전개(rollout) |
| L — Learning(학습) | 학습·환류 | 결과 분석, 교훈 정리, 다음 사이클로 반영 |

### 5-7. 개선 모델 비교 요약

| 구분 | 유형 | 진행 방식 | 강점 | 유의점 |
|------|------|-----------|------|--------|
| TMMi | 프로세스/모델 | 단계형(5레벨) | CMMI 정합, 공식 인증·명확한 로드맵 | 무겁고 전면 도입 부담 |
| TPI NEXT | 프로세스/모델 | 연속형(16영역) | 방법론 독립·유연, 부분 개선 용이 | 자체 표준화 강제력 약함 |
| CTP | 내용 기반 | 순서 비강제 | 경량·모든 SDLC 통합, 지표 벤치마킹 | 성숙도 로드맵 부재 |
| STEP | 내용 기반 | 순서 비강제 | 요구사항 기반·조기 테스트 강조 | 정량 성숙 측정 약함 |

---

## 6. 테스트 조직 KPI·품질 지표 설계

> 주의: 아래 지표 목록은 CTEL 원문 전용 목록이 아니라 ISTQB 테스트 관리 전반(CTAL-TM 포함)에서 통용되는 지표를 정리한 것. 지표는 목표(GQM)에서 도출해야 함.

### 6-1. 지표 설계 원칙(GQM 연계)

- 목표(Goal) → 질문(Question) → 지표(Metric) 순으로 도출. 지표를 먼저 고르지 않음
- 단일 지표 맹신 금지, 상충 지표를 함께 봄(예: 속도 vs 품질)
- 지표는 "행동을 왜곡"할 수 있음 → 조작 가능성(gaming) 고려해 설계

### 6-2. 대표 지표 범주

| 범주 | 예시 지표 | 보는 목적 |
|------|-----------|-----------|
| 진척(Progress) | 계획 대비 실행 케이스 수, 통과/실패/차단 비율 | 일정·완료 예측 |
| 결함(Defect) | 결함 검출 수·추세, 심각도 분포, 재개율(reopen rate) | 품질·안정성 판단 |
| 커버리지(Coverage) | 요구사항/코드/리스크 커버리지 | 테스트 충분성 |
| 효과성(Effectiveness) | DDE/DRE(결함 제거 효율), 유출 결함(escaped defects) | 프로세스가 결함을 잡는가 |
| 효율(Efficiency) | 결함당 소요·발견 비용, 테스트 자동화 비율·ROI | 비용 대비 성과 |
| 리스크(Risk) | 잔존 리스크 항목·커버 여부 | 릴리스 의사결정 |

- DRE(Defect Removal Efficiency) = 릴리스 전 발견 결함 / (릴리스 전 발견 + 릴리스 후 발견) — 프로세스 개선 성과의 대표 지표

---

## 7. 확인 불가 항목 (원문 대조 필요 / 미확인)

- TMMi 프로세스영역 개별 명칭: 레벨별 영역 구성은 TMMi Foundation 프레임워크에 근거해 정리했으나, 최신 릴리스(R 버전)에서 명칭·구성 미세 변경 가능. 공식 TMMi Framework 문서 대조 권장
- TPI NEXT 16개 핵심 영역 정확한 명칭·번역: 판본/번역에 따라 표기 차이. 정확한 공식 목록은 TPI NEXT 원서 대조 필요
- CTP 12개 임계 프로세스의 개별 명칭 12개 전체: 본 문서에서 개별 12개 명칭을 확정 나열하지 않음(할루시네이션 방지). 원문 대조 필요
- STEP 국면 명칭·세부 단계: Plan/Acquire/Measure 3국면으로 통용되나 세부 활동 명칭은 판본 차 — 원문 대조 필요
- TMMi "약 85% 달성" 진행 규칙: 널리 인용되나 공식 수치 기준인지 2차 출처 다수에 의존 — 공식 문서 확인 권장
- CTEL-TM-MTT(팀 관리) 세부 하위 학습목표: 공식 실러버스 원문(PDF)이 서버 접근 차단(403)으로 직접 대조 못 함. 상위 주제만 확인
- istqb.org 인증 페이지 원문: SM/OTM 상세 페이지가 403 차단으로 직접 확인 불가. 모듈 구성·현재 제공 상태는 istqb.guru 등 2차 출처와 검색 요약으로 교차 확인함

### 자가검증 메모(주의해서 볼 부분)
- 초기 PDF 자동요약이 TMMi 레벨4·5를 "Optimized/Innovative"로 잘못 냈으나, TMMi Foundation·복수 출처로 Measured(L4)·Optimization(L5) 확인해 정정함
- "2024년 Expert Level 단종"설은 오해. 실제 단종은 구버전 시험판(예: CTFL 3.1). Expert Level은 유지되며 7년 유효기간·갱신제

---

## 참고 출처

- ISTQB 공식 인증 안내: https://istqb.org/certifications/
- CTEL-TM Strategic(SM): https://istqb.org/certifications/certified-tester-expert-level-test-management-strategic-test-management-ctel-tm-sm/
- CTEL-TM Operational(OTM): https://istqb.org/certifications/certified-tester-expert-level-test-management-operational-test-management-ctel-tm-otm/
- CTEL-TM Managing the Test Team(MTT): https://istqb.org/certifications/certified-tester-expert-level-test-management-managing-the-test-team-ctel-tm-mtt/
- CTEL-ITP Assessing the Test Process(ATP): https://istqb.org/certifications/certified-tester-expert-level-assessing-test-processes-ctel-itp-atp/
- CTEL-ITP Implementing Test Process Improvement(ITPI): https://istqb.org/sdm_categories/certified-tester-expert-level-implementing-test-process-improvement-ctel-itp-itpi/
- CTEL Improving the Testing Process 실러버스(PDF): https://istqb.org/wp-content/uploads/2024/11/ISTQB-CTEL-ITP_Syllabus_v1.0_2011.pdf (미러: https://sjsi.org/wp-content/uploads/2013/11/ISTQB_EL_ImprovingTheTestProcess_EN.pdf )
- Expert Level 갱신 정책(PDF): https://istqb.org/wp-content/uploads/2024/11/ISTQB_EL_Certification_Extension_Policy_v1.1.pdf
- 개선 모델 요약(TMMi/TPI NEXT/CTP/STEP): https://tryqa.com/software-testing-process-improvement-models-tmmi-tpi-next-ctp-step/
- TMMi Foundation 모델: https://www.tmmi.org/tmmi-model/ , TMMi Framework(PDF): https://tmmi.org/tm6/wp-content/uploads/2018/11/TMMi-Framework-R1-2.pdf
- ISTQB 인증 로드맵(2차): https://www.istqb.guru/istqb-certification-levels-roadmap/
- ASTQB Expert Test Management 자료: https://astqb.org/istqb-expert-level-test-management-syllabus/
