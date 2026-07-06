#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sync_readme.py — README의 스킬 목록 표를 marketplace.json 기준으로 자동 생성한다.

하는 일:
  .claude-plugin/marketplace.json 의 plugins 목록을 읽어,
  README.md 의 <!-- SKILLS:START ... --> ~ <!-- SKILLS:END --> 사이 표를 다시 채운다.
  (표는 손으로 고치지 말고 이 스크립트로만 갱신한다.)

사용법:
  python _tools/sync_readme.py [저장소_경로]
  - 경로를 생략하면 이 스크립트의 상위(저장소 루트)를 자동으로 쓴다.
"""

import json
import re
import sys
from pathlib import Path

# 윈도우 콘솔(cp949)에서도 한글·특수문자 출력이 깨지지 않게 utf-8로 맞춘다.
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# 설명 컬럼은 " — "(공백+em대시+공백) 앞부분만 짧게 쓴다. 없으면 전체를 쓴다.
SEP = " — "

START = re.compile(r"(<!--\s*SKILLS:START.*?-->)", re.DOTALL)
END = "<!-- SKILLS:END -->"


def short_desc(desc: str) -> str:
    return desc.split(SEP, 1)[0].strip()


def build_table(marketplace: dict) -> str:
    mkt_name = marketplace.get("name", "").strip()
    rows = ["| 스킬 | 설명 | 설치 명령 |", "|---|---|---|"]
    for p in marketplace.get("plugins", []):
        name = p.get("name", "").strip()
        desc = short_desc(p.get("description", ""))
        install = f"`/plugin install {name}@{mkt_name}`"
        rows.append(f"| **{name}** | {desc} | {install} |")
    return "\n".join(rows)


def main() -> int:
    repo = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent
    mkt_path = repo / ".claude-plugin" / "marketplace.json"
    readme_path = repo / "README.md"

    if not mkt_path.exists():
        print(f"[오류] marketplace.json을 찾을 수 없어요: {mkt_path}")
        return 1
    if not readme_path.exists():
        print(f"[오류] README.md를 찾을 수 없어요: {readme_path}")
        return 1

    marketplace = json.loads(mkt_path.read_text(encoding="utf-8"))
    readme = readme_path.read_text(encoding="utf-8")

    start_m = START.search(readme)
    end_i = readme.find(END)
    if not start_m or end_i == -1 or end_i < start_m.end():
        print("[오류] README.md에 <!-- SKILLS:START --> ~ <!-- SKILLS:END --> 표식이 없어요.")
        return 1

    table = build_table(marketplace)
    new_readme = (
        readme[: start_m.end()]
        + "\n"
        + table
        + "\n"
        + readme[end_i:]
    )

    if new_readme == readme:
        print("변경 없음: README 표가 이미 최신이에요.")
        return 0

    readme_path.write_text(new_readme, encoding="utf-8", newline="\n")
    count = len(marketplace.get("plugins", []))
    print(f"README 표 갱신 완료: 플러그인 {count}개 반영.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
