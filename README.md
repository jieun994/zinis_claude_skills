# zinis_claude_skills

신지은의 **Claude 스킬 모음** 마켓플레이스입니다. 마켓플레이스를 한 번만 등록하면
필요한 스킬을 골라 설치할 수 있어요.

## 설치 방법

### 데스크톱 앱 (Claude Code Desktop)
채팅창에 `/plugin` 을 입력하면 안 돼요. **GUI로** 설치합니다:
1. 메시지 입력창 옆 **`+` 버튼** → **Plugins**
2. 마켓플레이스 추가에 `jieun994/zinis_claude_skills` 등록 → `wbs-generator` 설치

GUI에서 마켓플레이스 추가가 안 보이면, `.claude/settings.json` 에 아래를 넣으세요:
```json
{
  "extraKnownMarketplaces": {
    "zinis_claude_skills": {
      "source": { "source": "github", "repo": "jieun994/zinis_claude_skills" }
    }
  }
}
```

### 터미널 (CLI)
```
/plugin marketplace add jieun994/zinis_claude_skills
/plugin install wbs-generator@zinis_claude_skills
```
(또는 `claude plugin marketplace add jieun994/zinis_claude_skills`)

## 업데이트 (최신 버전 받기)

- 데스크톱: `+` → Plugins 에서 업데이트
- 터미널: `/plugin marketplace update zinis_claude_skills`

> 자동 업데이트는 없어요. 새 버전 안내를 받으면 한 번 갱신하면 됩니다.

## 스킬 목록

| 스킬 | 설명 | 설치 명령 |
|---|---|---|
| **wbs-generator** | 프로젝트 WBS(작업분해구조)를 엑셀로 자동 생성 | `/plugin install wbs-generator@zinis_claude_skills` |

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
