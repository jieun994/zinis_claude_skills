# zinis_claude_skills

신지은의 **Claude 스킬 모음** 마켓플레이스입니다. 마켓플레이스를 한 번만 등록하면
필요한 스킬을 골라 설치할 수 있어요.

## 설치 방법

> 한 번만 설정하면, 그다음부터는 이 저장소의 스킬을 자유롭게 골라 쓸 수 있어요.
> 본인이 쓰는 환경(데스크톱 앱 / 터미널)에 맞는 쪽만 따라 하면 됩니다.

### 🖥️ 데스크톱 앱 (Claude Code Desktop) — 대부분 이 경우예요

> ⚠️ 데스크톱에서는 채팅창에 `/plugin` 이라고 쳐도 안 돼요. **메뉴(GUI)로** 설치합니다.

**1단계. 마켓플레이스 등록 (처음 한 번만)**

설정 파일에 아래를 추가합니다. 파일 위치: `C:\Users\<내계정>\.claude\settings.json`
(Mac은 `~/.claude/settings.json`)

```json
{
  "extraKnownMarketplaces": {
    "zinis_claude_skills": {
      "source": { "source": "github", "repo": "jieun994/zinis_claude_skills" }
    }
  }
}
```
- 파일에 이미 다른 내용이 있으면, `extraKnownMarketplaces` 안에 위 항목만 **추가**하세요(전체를 지우지 말고).

**2단계. Claude Code를 껐다 켜기 (재시작)**
- 설정 파일은 시작할 때 읽혀요. 그래서 저장 후 **앱을 완전히 종료했다가 다시 실행**해야 적용됩니다.

**3단계. 디렉터리에서 설치**
1. **디렉터리(Directory)** 열기 → 왼쪽 **플러그인** 선택
2. 상단 탭에서 **`코드`** 클릭 (Anthropic·파트너·**코드** 중 마지막)
3. 검색창에 **`wbs`** 입력 (또는 목록에서 `wbs-generator` 찾기)
4. **wbs-generator** 카드의 **`+` 버튼** 클릭 → 설치 완료!

> 안 보이면? → 새로고침(🔄) 누르거나, 2단계(재시작)를 다시 해보세요.

### ⌨️ 터미널 (Claude Code CLI)

터미널에서 `claude` 로 실행한 상태라면 채팅에 그대로 입력하면 돼요:
```
/plugin marketplace add jieun994/zinis_claude_skills
/plugin install wbs-generator@zinis_claude_skills
```

### 📁 마켓플레이스 없이 — 스킬 폴더 직접 복사 (가장 간단)

GitHub·마켓플레이스 없이, **`skills/wbs-generator` 폴더만** 받아서 넣으면 돼요.

1. 이 저장소에서 **`skills/wbs-generator`** 폴더를 받기 (다운로드 또는 zip)
2. 아래 위치에 폴더째로 넣기:
   - 모든 프로젝트에서 쓰려면 → `~/.claude/skills/wbs-generator/`
     (Windows: `C:\Users\<내계정>\.claude\skills\wbs-generator\`)
   - 특정 프로젝트에서만 → 그 프로젝트의 `.claude/skills/wbs-generator/`
3. Claude Code 재시작 → "WBS 만들어줘" 작동

> 가장 쉬운 대신, **업데이트는 폴더를 다시 받아 덮어써야** 해요(자동 아님).
> 여러 명이 계속 최신을 받아야 하면 위의 **마켓플레이스 방식**이 편합니다.

### 설치 후

Claude에게 **"WBS 만들어줘"** 라고 하면 스킬이 작동합니다. 🎉

## 업데이트 (새 버전이 나왔을 때)

- **데스크톱:** 디렉터리 → 플러그인 → 코드 탭 → wbs-generator → 업데이트
- **터미널:** `/plugin marketplace update zinis_claude_skills`

> 자동 업데이트는 없어요. "업데이트됐어요" 안내를 받으면 위 방법으로 **한 번만** 갱신하면 최신이 됩니다.

## 스킬 목록

<!-- SKILLS:START — 이 표는 _tools/sync_readme.py가 marketplace.json에서 자동 생성합니다. 직접 수정하지 마세요. -->
| 스킬 | 설명 | 설치 명령 |
|---|---|---|
| **wbs-generator** | 프로젝트 WBS를 엑셀로 자동 생성 | `/plugin install wbs-generator@zinis_claude_skills` |
| **ia-builder** | 서비스 메뉴 구조(사이트맵/IA)와 기능정의서를 엑셀로 자동 생성 | `/plugin install ia-builder@zinis_claude_skills` |
| **flowchart-defense** | 고객이 준 흐름도(PPT·이미지·PDF·mermaid)를 검토하고 방어설계로 보강 | `/plugin install flowchart-defense@zinis_claude_skills` |
<!-- SKILLS:END -->

각 스킬의 자세한 사용법은 해당 폴더의 README를 보세요.
(예: [skills/wbs-generator/README.md](skills/wbs-generator/README.md))

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
2. `.claude-plugin/marketplace.json` 의 `plugins` 에 한 줄 등록
   (`name`, `source: "./skills/<새스킬>"`)
3. commit & push → 사용자는 `/plugin install <새스킬>@zinis_claude_skills` 로 설치
