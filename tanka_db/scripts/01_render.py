#!/usr/bin/env python3
"""
01_render.py — PDF を 1 ページ 1 画像へレンダリングする。

設計書 §3「01_render.py」対応。
- DocuWorks 由来でテキスト層ゼロ・全ページ 90°回転のため、pdftoppm で
  正立ラスタ化する（poppler-utils が必要）。
- DPI は精度検証のため 150/200 を切替可能（既定 200）。

使い方:
  python 01_render.py                      # input_pdfs/ の全 PDF を処理
  python 01_render.py --pdf path/to/x.pdf  # 1 ファイルだけ
  python 01_render.py --dpi 150            # DPI 変更
  python 01_render.py --first 1 --last 3   # ページ範囲を限定（検証用）

出力: work/{work_id}/pages/p{NNN}.jpg
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

import common


def check_poppler() -> None:
    if shutil.which("pdftoppm") is None:
        sys.exit(
            "エラー: pdftoppm が見つかりません。poppler-utils をインストールしてください。\n"
            "  Debian/Ubuntu: sudo apt-get install poppler-utils\n"
            "  macOS(brew):    brew install poppler"
        )


def render_pdf(pdf_path: Path, dpi: int, first: int | None, last: int | None) -> Path:
    work_id = common.work_id_from_pdf(pdf_path)
    out_dir = common.pages_dir(work_id)
    out_dir.mkdir(parents=True, exist_ok=True)

    # pdftoppm は出力プレフィックスに -NN を付ける。桁揃えは -sep なしのため
    # 後段で p{NNN}.jpg へ寄せる。ここでは prefix 'page' で出力してからリネーム。
    prefix = out_dir / "page"
    cmd = ["pdftoppm", "-jpeg", "-r", str(dpi)]
    if first is not None:
        cmd += ["-f", str(first)]
    if last is not None:
        cmd += ["-l", str(last)]
    cmd += [str(pdf_path), str(prefix)]

    print(f"[render] {work_id}  dpi={dpi}  -> {out_dir}")
    subprocess.run(cmd, check=True)

    # page-1.jpg / page-01.jpg 等を p{NNN}.jpg に正規化。
    renamed = 0
    for f in sorted(out_dir.glob("page-*.jpg")):
        num = int(f.stem.split("-")[-1])
        target = out_dir / f"p{num:03d}.jpg"
        if f != target:
            f.replace(target)
        renamed += 1
    print(f"[render] {work_id}  {renamed} ページを出力しました。")
    return out_dir


def main() -> None:
    ap = argparse.ArgumentParser(description="PDF -> ページ画像")
    ap.add_argument("--pdf", help="単一 PDF のパス。省略時は input_pdfs/ の全 PDF")
    ap.add_argument("--input-dir", default=str(common.INPUT_PDFS))
    ap.add_argument("--dpi", type=int, default=200, help="解像度(既定200。検証で150と比較)")
    ap.add_argument("--first", type=int, default=None, help="開始ページ(1始まり)")
    ap.add_argument("--last", type=int, default=None, help="終了ページ")
    args = ap.parse_args()

    check_poppler()

    if args.pdf:
        pdfs = [Path(args.pdf)]
    else:
        pdfs = sorted(Path(args.input_dir).glob("*.pdf"))

    if not pdfs:
        sys.exit(f"PDF が見つかりません: {args.pdf or args.input_dir}")

    for pdf in pdfs:
        render_pdf(pdf, args.dpi, args.first, args.last)


if __name__ == "__main__":
    main()
