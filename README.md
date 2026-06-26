# wbs-generator

프로젝트 **WBS(작업 분해 구조)를 엑셀로 자동 생성**하는 Claude 스킬 마켓플레이스입니다.

대화로 프로젝트를 설명하면 작업을 단계별(L1~L3)로 나누고, 일정·담당자·간트차트·
진척률·공휴일까지 채운 엑셀을 만들어 줍니다.

## 설치 방법

Claude Code에서 아래를 실행하세요. (private 저장소라 **초대받은 계정으로 GitHub에
로그인된 상태**여야 합니다)

```
/plugin marketplace add jieun994/wbs-generator
/plugin install wbs-generator@wbs-generator
```

설치 후 Claude에게 **"WBS 만들어줘"** 라고 하면 됩니다.

## 들어있는 스킬

| 스킬 | 설명 |
|---|---|
| **wbs-generator** | 프로젝트 WBS를 엑셀로 자동 생성 (레벨별 작업·일정·간트·진척률·공휴일) |

자세한 사용법은 [skills/wbs-generator/README.md](skills/wbs-generator/README.md) 참고.

## 필요 환경

- Python 패키지: `pip install openpyxl holidays`

## 폴더 구성

```
.claude-plugin/
  marketplace.json   ← 마켓플레이스 목록
  plugin.json        ← 플러그인 정보
skills/
  wbs-generator/     ← 실제 스킬 (SKILL.md · README.md · scripts · template · samples)
```
