# claude-skills

신지은의 **Claude 스킬 모음** 마켓플레이스입니다. 마켓플레이스를 한 번만 등록하면
필요한 스킬을 골라 설치할 수 있어요.

## 설치 방법

Claude Code에서:

```
# 1) 마켓플레이스 등록 (처음 한 번만)
/plugin marketplace add jieun994/claude-skills

# 2) 원하는 스킬 설치
/plugin install wbs-generator@claude-skills
```

## 업데이트 (최신 버전 받기)

```
/plugin marketplace update claude-skills
```

> 자동 업데이트는 없어요. 새 버전 안내를 받으면 위 명령을 한 번 실행하면 됩니다.

## 스킬 목록

| 스킬 | 설명 | 설치 명령 |
|---|---|---|
| **wbs-generator** | 프로젝트 WBS(작업분해구조)를 엑셀로 자동 생성 | `/plugin install wbs-generator@claude-skills` |

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
3. commit & push → 사용자는 `/plugin install <새스킬>@claude-skills` 로 설치
