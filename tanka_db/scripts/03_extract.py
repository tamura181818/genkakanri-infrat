#!/usr/bin/env python3
"""
03_extract.py — ページ画像から 種別判定＋構造化抽出（ビジョン LLM）。

設計書 §3「02_classify / 03_extract」の統合方針を実装。
各ページを 1 リクエストで「種別判定 + 対象なら明細抽出」する。鏡ページは
別プロンプトでメタ情報を抽出し、work メタとして保存する。

使い方:
  python 03_extract.py                       # work/ 配下の全 work_id を処理
  python 03_extract.py --work-id 01_R7国道…   # 1 工事だけ
  python 03_extract.py --model gemini-2.5-flash-lite

出力:
  work/{work_id}/extracted.jsonl   … 明細行（ページ単位・元の値のまま）
  work/{work_id}/meta.json         … 鏡から抽出したメタ情報
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import common
import gemini_client
import prompts


def process_work(work_id: str, model_name: str, overwrite: bool) -> None:
    pdir = common.pages_dir(work_id)
    pages = sorted(pdir.glob("p*.jpg"))
    if not pages:
        print(f"[extract] {work_id}: ページ画像なし。先に 01_render を実行してください。")
        return

    out_jsonl = common.work_dir(work_id) / "extracted.jsonl"
    meta_path = common.work_dir(work_id) / "meta.json"
    if out_jsonl.exists() and not overwrite:
        print(f"[extract] {work_id}: {out_jsonl.name} が既存（--overwrite で再抽出）。skip")
        return

    all_rows: list[dict] = []
    meta: dict = {}

    for page_img in pages:
        page_no = int(page_img.stem.lstrip("p"))
        # 1 ページ目は鏡として先にメタ抽出を試みる。
        if page_no == 1:
            try:
                meta = gemini_client.call_vision_json(
                    str(page_img), prompts.KAGAMI_PROMPT, prompts.KAGAMI_SCHEMA, model_name
                )
                meta["work_id"] = work_id
                print(f"[extract] {work_id} p{page_no:03d}: 鏡メタ抽出 OK")
            except Exception as e:  # 鏡でない/読めない場合は明細抽出にフォールバック
                print(f"[extract] {work_id} p{page_no:03d}: 鏡抽出失敗({e}) -> 明細抽出へ")

        try:
            res = gemini_client.call_vision_json(
                str(page_img), prompts.EXTRACT_PROMPT, prompts.EXTRACT_SCHEMA, model_name
            )
        except gemini_client.GeminiUnavailable as e:
            sys.exit(f"[extract] 中断: {e}")
        except Exception as e:
            print(f"[extract] {work_id} p{page_no:03d}: 抽出失敗 {e}")
            continue

        page_type = res.get("page_type", "その他")
        page_go = res.get("摘要号")
        rows = res.get("rows") or []
        for r in rows:
            r["ページ"] = page_no
            r["表種別"] = page_type
            if page_go and not r.get("摘要号"):
                r["摘要号"] = page_go
            all_rows.append(r)
        if rows:
            print(f"[extract] {work_id} p{page_no:03d}: {page_type} / {len(rows)}行")

    common.write_jsonl(out_jsonl, all_rows)
    if meta:
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[extract] {work_id}: 合計 {len(all_rows)}行 -> {out_jsonl}")


def main() -> None:
    ap = argparse.ArgumentParser(description="ページ画像 -> 構造化抽出(ビジョンLLM)")
    ap.add_argument("--work-id", help="単一 work_id。省略時は work/ 配下すべて")
    ap.add_argument("--model", default=gemini_client.DEFAULT_MODEL,
                    help="gemini-2.5-flash / gemini-2.5-flash-lite")
    ap.add_argument("--overwrite", action="store_true", help="既存 extracted.jsonl を再抽出")
    args = ap.parse_args()

    if args.work_id:
        work_ids = [args.work_id]
    else:
        work_ids = [p.name for p in sorted(common.WORK.iterdir())
                    if p.is_dir() and (p / "pages").exists()]
    if not work_ids:
        sys.exit("処理対象の work_id がありません。先に 01_render を実行してください。")

    for wid in work_ids:
        process_work(wid, args.model, args.overwrite)


if __name__ == "__main__":
    main()
