# zinis_claude_skills

신지은의 **Claude 스킬 모음** 마켓플레이스입니다. 한 번만 등록하면 필요한 스킬을 골라 설치할 수 있습니다.

## 설치 방법

- 한 번만 설정하면, 이후 이 저장소의 스킬을 자유롭게 골라 사용
- 본인 환경(데스크톱 앱 / 터미널)에 맞는 쪽만 진행

### 데스크톱 앱 (Claude Code Desktop) — 대부분 이 경우

- 데스크톱에서는 채팅창에 `/plugin` 을 쳐도 안 됨 → **메뉴(GUI)로** 설치

**1단계. 마켓플레이스 등록 (처음 한 번만)**

- 설정 파일에 아래 추가
- 파일 위치: `C:\Users\<내계정>\.claude\settings.json` (Mac은 `~/.claude/settings.json`)

```json
{
  "extraKnownMarketplaces": {
    "zinis_claude_skills": {
      "source": { "source": "github", "repo": "jieun994/zinis_claude_skills" }
    }
  }
}
```

- 이미 다른 내용이 있으면 `extraKnownMarketplaces` 안에 위 항목만 **추가** (전체 삭제 금지)

**2단계. Claude Code 재시작**

- 설정 파일은 시작할 때 읽힘 → 저장 후 **앱 완전 종료 후 재실행**해야 적용

**3단계. 디렉터리에서 설치**

1. **디렉터리(Directory)** 열기 → 왼쪽 **플러그인** 선택
2. 상단 탭에서 **`코드`** 클릭 (Anthropic·파트너·**코드** 중 마지막)
3. 검색창에 **`wbs`** 입력 (또는 목록에서 `wbs-generator` 찾기)
4. **wbs-generator** 카드의 **`+` 버튼** 클릭 → 설치 완료

- 안 보이면: 새로고침 누르거나 2단계(재시작) 다시 시도

### 터미널 (Claude Code CLI)

- 터미널에서 `claude` 로 실행한 상태면 채팅에 그대로 입력

```
/plugin marketplace add jieun994/zinis_claude_skills
/plugin install wbs-generator@zinis_claude_skills
```

### 마켓플레이스 없이 — 스킬 폴더 직접 복사 (가장 간단)

- GitHub·마켓플레이스 없이 **`skills/wbs-generator` 폴더만** 받아서 넣으면 됨

1. 이 저장소에서 **`skills/wbs-generator`** 폴더 받기 (다운로드 또는 zip)
2. 아래 위치에 폴더째로 넣기
   - 모든 프로젝트에서 사용 → `~/.claude/skills/wbs-generator/` (Windows: `C:\Users\<내계정>\.claude\skills\wbs-generator\`)
   - 특정 프로젝트에서만 → 그 프로젝트의 `.claude/skills/wbs-generator/`
3. Claude Code 재시작 → "WBS 만들어줘" 작동

- 가장 쉽지만 **업데이트는 폴더를 다시 받아 덮어써야 함**(자동 아님)
- 여러 명이 계속 최신을 받아야 하면 위 **마켓플레이스 방식**이 편함

### 설치 후

- Claude에게 **"WBS 만들어줘"** 라고 하면 스킬 작동

## 업데이트 (새 버전이 나왔을 때)

- **데스크톱:** 디렉터리 → 플러그인 → 코드 탭 → wbs-generator → 업데이트
- **터미널:** `/plugin marketplace update zinis_claude_skills`
- 자동 업데이트 없음 → "업데이트됐어요" 안내를 받으면 위 방법으로 **한 번만** 갱신

## 스킬 목록

<!-- SKILLS:START — 이 표는 _tools/sync_readme.py가 marketplace.json에서 자동 생성합니다. 직접 수정하지 마세요. -->
| 스킬 | 설명 | 설치 명령 |
|---|---|---|
| **wbs-generator** | 프로젝트 WBS를 엑셀로 자동 생성 | `/plugin install wbs-generator@zinis_claude_skills` |
| **ia-builder** | 서비스 메뉴 구조(사이트맵/IA)와 기능정의서를 엑셀로 자동 생성 | `/plugin install ia-builder@zinis_claude_skills` |
| **flowchart-defense** | 고객이 준 흐름도(PPT·이미지·PDF·mermaid)를 검토하고 방어설계로 보강 | `/plugin install flowchart-defense@zinis_claude_skills` |
| **weekly-report** | 주간업무 보고서(PPT)를 기존 양식 그대로 자동 작성 | `/plugin install weekly-report@zinis_claude_skills` |
| **dev-reading** | 개발자료(API 명세·DB 스키마/ERD·기술문서)를 비개발 기획자가 이해하도록 해설 | `/plugin install dev-reading@zinis_claude_skills` |
| **requirements-generator** | 요구사항 정의서를 엑셀로 자동 생성 | `/plugin install requirements-generator@zinis_claude_skills` |
| **meeting-minutes** | 녹취록·메모로 표준 엑셀 회의록 자동 작성 | `/plugin install meeting-minutes@zinis_claude_skills` |
| **jieun** | 신지은님처럼 일하는 분신 에이전트 모음(나답게) | `/plugin install jieun@zinis_claude_skills` |
<!-- SKILLS:END -->

- 각 스킬 자세한 사용법은 해당 폴더의 README 참고 (예: [skills/wbs-generator/README.md](skills/wbs-generator/README.md))

## 폴더 구성

```
.claude-plugin/
  marketplace.json     ← 스킬 카탈로그 (여기에 스킬 등록)
skills/
  wbs-generator/       ← 각 스킬 (자체 .claude-plugin/plugin.json + SKILL.md 포함)
  …                    ← 새 스킬은 폴더로 추가
```

## 새 스킬 추가 방법 (관리자용)

1. `skills/<새스킬>/` 폴더 추가 (SKILL.md + `.claude-plugin/plugin.json` 포함)
2. `.claude-plugin/marketplace.json` 의 `plugins` 에 한 줄 등록 (`name`, `source: "./skills/<새스킬>"`)
3. commit & push → 사용자는 `/plugin install <새스킬>@zinis_claude_skills` 로 설치
