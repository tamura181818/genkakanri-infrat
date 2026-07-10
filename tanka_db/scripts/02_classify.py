#!/usr/bin/env python3
"""
02_classify.py — ページ種別のみを判定する（任意・単体運用）。

設計書 §3「02_classify.py」対応。既定パイプラインでは 03_extract が
種別判定＋抽出を統合して行うため本スクリプトは必須ではないが、
- 抽出前にページ構成を俯瞰したい
- 種別判定だけのコスト/精度を単体で測りたい
という用途向けに残す。

使い方:
  python 02_classify.py --work-id 01_R7国道…
出力: work/{work_id}/page_types.json  （{ "p001": "鏡", ... }）
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import common
import gemini_client
import prompts

CLASSIFY_PROMPT = (
    "渡された画像は工事設計書の1ページです。種別を次から1つ選んで JSON で返してください: "
    + " / ".join(prompts.PAGE_TYPES)
    + "。基準は 鏡=メタ情報の1ページ目 / 設計内訳書=工事区分・工種の内訳表 / "
    "内訳書=一式当たり内訳書(内-○号) / 単価表=1次2次単価表(単-○号) / その他=表紙図面等。"
    '出力は {"page_type": "..."} の JSON のみ。'
)
CLASSIFY_SCHEMA = {
    "type": "object",
    "properties": {"page_type": {"type": "string", "enum": prompts.PAGE_TYPES}},
    "required": ["page_type"],
}


def classify_work(work_id: str, model_name: str) -> None:
    pages = sorted(common.pages_dir(work_id).glob("p*.jpg"))
    if not pages:
        print(f"[classify] {work_id}: ページ画像なし。先に 01_render を実行してください。")
        return
    result: dict[str, str] = {}
    for page_img in pages:
        try:
            res = gemini_client.call_vision_json(
                str(page_img), CLASSIFY_PROMPT, CLASSIFY_SCHEMA, model_name
            )
        except gemini_client.GeminiUnavailable as e:
            sys.exit(f"[classify] 中断: {e}")
        result[page_img.stem] = res.get("page_type", "その他")
        print(f"[classify] {work_id} {page_img.stem}: {result[page_img.stem]}")
    out = common.work_dir(work_id) / "page_types.json"
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[classify] {work_id}: -> {out}")


def main() -> None:
    ap = argparse.ArgumentParser(description="ページ種別判定(単体)")
    ap.add_argument("--work-id", help="単一 work_id。省略時は work/ 配下すべて")
    ap.add_argument("--model", default=gemini_client.DEFAULT_MODEL)
    args = ap.parse_args()

    if args.work_id:
        work_ids = [args.work_id]
    else:
        work_ids = [p.name for p in sorted(common.WORK.iterdir())
                    if p.is_dir() and (p / "pages").exists()]
    if not work_ids:
        sys.exit("処理対象の work_id がありません。")
    for wid in work_ids:
        classify_work(wid, args.model)


if __name__ == "__main__":
    main()
