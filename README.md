# zinis_claude_skills — 기획 업무용 Claude 스킬 모음

대화만으로 기획 산출물(WBS·IA·플로우차트·주간보고 등)을 자동 생성하는 Claude 스킬 마켓플레이스입니다.

## 설치 방법

Claude Code에서 아래 실행:

```
/plugin marketplace add jieun994/zinis_claude_skills
/plugin install <스킬이름>@zinis_claude_skills
```

- 업데이트: `/plugin marketplace update zinis_claude_skills`

## 들어있는 스킬

| 스킬                    | 설명                                                    |
| --------------------- | ----------------------------------------------------- |
| **wbs-generator**     | 프로젝트 WBS를 엑셀로 자동 생성 (레벨별 작업·일정·간트·진척률·공휴일)            |
| **ia-builder**        | 사이트맵·기능정의서(IA)를 엑셀로 자동 생성 (Depth 1~5, Front/Admin 분리) |
| **flowchart-defense** | 고객 흐름도 검토·방어설계 (패턴 1~8, MD 리포트 + 워드 전달본)              |
| **weekly-report**     | 주간보고서 PPT 자동 작성 (W1/W2, 진척률 자동 계산)                    |
| **dev-reading**       | 개발자료(API·ERD·기술문서)를 기획자 눈높이로 해설                       |

- 분신 에이전트: **jieun** (메일 쓰는 나 · 기획하는 나 · 개발 통역하는 나)
- 각 스킬 자세한 사용법은 `skills/<스킬>/README.md` 참고

## 필요 환경

- Python 패키지 (스킬별로 다름)
  - wbs-generator: `pip install openpyxl holidays`
  - ia-builder: `pip install openpyxl`
  - flowchart-defense: 워드 이미지 변환 시 `mermaid-cli` (선택)

## 폴더 구성

```
skills/
  <스킬이름>/          ← 실제 스킬 (SKILL.md · README.md · samples 등)
personas/
  <이름>/              ← 분신 에이전트 플러그인
.claude-plugin/
  marketplace.json     ← 마켓플레이스 목록
```
