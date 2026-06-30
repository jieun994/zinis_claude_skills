#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render_slides.py — PPTX 슬라이드를 PNG 이미지로 렌더(화살표 연결 비전 판독용).

extract_pptx.py는 텍스트·좌표만 주므로, 화살표 연결 복원은 슬라이드를 이미지로
렌더해 비전으로 봐야 한다. 이 스크립트가 그 이미지를 만든다.

우선순위:
  1) PowerPoint COM (Windows + MS PowerPoint 설치)  → 슬라이드별 PNG export (가장 정확)
  2) LibreOffice(soffice)                          → 전체를 PDF로 변환(폴백, 페이지=슬라이드)

사용법:
  python render_slides.py <파일.pptx> [--out DIR] [--slides 1,2,6] [--width 2200] [--height 2900]
    --out     출력 폴더 (기본: <pptx>_slides)
    --slides  렌더할 슬라이드 번호(1부터), 콤마구분. 생략 시 전체
    --width   PNG 가로 px (기본 2200)
    --height  PNG 세로 px (기본 2900)

판독 팁:
  - 슬라이드가 넓거나 세로로 길면 가장자리가 잘릴 수 있다. 출력 로그의 슬라이드 크기를
    확인하고, 잘리면 --width/--height를 키우거나 이미지 편집으로 영역을 크롭해 읽는다.
  - 좌표가 캔버스 폭을 넘는 도형은 슬라이드 밖일 수 있으나, '바깥'이라 단정하지 말고
    렌더 이미지로 실제 연결을 확인한다.
"""
import os
import sys
import argparse


def parse_args():
    ap = argparse.ArgumentParser(description="PPTX 슬라이드 → PNG 렌더")
    ap.add_argument("pptx", help="입력 .pptx 경로")
    ap.add_argument("--out", default=None, help="출력 폴더")
    ap.add_argument("--slides", default=None, help="슬라이드 번호(1부터), 콤마구분. 예: 1,2,6")
    ap.add_argument("--width", type=int, default=2200, help="PNG 가로 px")
    ap.add_argument("--height", type=int, default=2900, help="PNG 세로 px")
    return ap.parse_args()


def render_with_powerpoint(pptx, out_dir, slides, width, height):
    """Windows PowerPoint COM으로 슬라이드별 PNG export. 성공 시 True."""
    try:
        try:
            import win32com.client as com  # pywin32
            ppt = com.Dispatch("PowerPoint.Application")
        except Exception:
            import comtypes.client as com  # comtypes 폴백
            ppt = com.CreateObject("PowerPoint.Application")
    except Exception as e:
        print(f"[PowerPoint COM 사용 불가] {e}")
        return False

    pres = None
    try:
        pres = ppt.Presentations.Open(os.path.abspath(pptx), WithWindow=False)
    except Exception:
        # WithWindow 키워드가 안 먹는 바인딩(positional): ReadOnly, Untitled, WithWindow
        try:
            pres = ppt.Presentations.Open(os.path.abspath(pptx), True, False, False)
        except Exception as e:
            print(f"[열기 실패] {e}")
            try:
                ppt.Quit()
            except Exception:
                pass
            return False

    try:
        total = pres.Slides.Count
        try:
            sw = round(float(pres.PageSetup.SlideWidth))
            sh = round(float(pres.PageSetup.SlideHeight))
            print(f"슬라이드 크기(pt): {sw} x {sh}  (가로>세로면 와이드, 반대면 세로형)")
        except Exception:
            pass
        targets = slides if slides else range(1, total + 1)
        done = []
        for n in targets:
            if n < 1 or n > total:
                print(f"  - slide {n}: 범위 밖(전체 {total}) 건너뜀")
                continue
            out = os.path.join(out_dir, f"slide{n}.png")
            pres.Slides.Item(n).Export(out, "PNG", width, height)
            done.append(out)
            print(f"  - slide {n} → {out}")
        return True if done else False
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        try:
            ppt.Quit()
        except Exception:
            pass


def render_with_libreoffice(pptx, out_dir):
    """soffice로 PDF 변환(폴백). 페이지=슬라이드. PNG가 필요하면 PDF를 비전으로 읽거나 별도 변환."""
    import shutil
    import subprocess

    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if not soffice:
        print("[LibreOffice 없음] soffice를 찾지 못함.")
        return False
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, os.path.abspath(pptx)],
            check=True,
        )
        base = os.path.splitext(os.path.basename(pptx))[0]
        pdf = os.path.join(out_dir, base + ".pdf")
        print(f"  - PDF 생성 → {pdf}")
        print("    (PDF 페이지를 비전으로 읽거나, pdftoppm/pymupdf로 PNG 변환해 사용)")
        return os.path.exists(pdf)
    except Exception as e:
        print(f"[LibreOffice 변환 실패] {e}")
        return False


def main():
    args = parse_args()
    if not os.path.exists(args.pptx):
        print(f"파일 없음: {args.pptx}")
        sys.exit(1)

    out_dir = args.out or (os.path.splitext(args.pptx)[0] + "_slides")
    os.makedirs(out_dir, exist_ok=True)

    slides = None
    if args.slides:
        slides = [int(x) for x in args.slides.replace(" ", "").split(",") if x]

    print(f"입력: {args.pptx}")
    print(f"출력 폴더: {out_dir}")

    if render_with_powerpoint(args.pptx, out_dir, slides, args.width, args.height):
        print("완료(PowerPoint COM).")
        return
    print("PowerPoint COM 실패 → LibreOffice 폴백 시도")
    if render_with_libreoffice(args.pptx, out_dir):
        print("완료(LibreOffice PDF).")
        return
    print(
        "렌더 실패. 다음 중 하나가 필요합니다:\n"
        "  · Windows + MS PowerPoint (+ pywin32 또는 comtypes)\n"
        "  · 또는 LibreOffice(soffice)\n"
        "수동 대안: PowerPoint에서 '다른 이름으로 저장 → PNG'로 슬라이드를 내보내 비전 판독."
    )
    sys.exit(2)


if __name__ == "__main__":
    main()
