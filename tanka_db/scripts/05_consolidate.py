#!/usr/bin/env python3
"""
05_consolidate.py — 全工事の validated.jsonl を統合し CSV を出力。

設計書 §3「05_consolidate.py」対応。
- output/単価一覧.csv      … is_kept=True の DB 本体（UTF-8 BOM なし）
- output/review_flags.csv  … 要レビュー・検算超過・除外行（監査/辞書育成用）

重複方針(§3):
  ・同一工事内の重複（同一号が複数箇所から参照され同じ明細が重複）→ 除去。
  ・工事・単価適用年月が異なる同一品目 → 別レコードとして残す
    （時系列・事務所間の単価比較に使うため）。
"""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

import common

# §2 データモデルの出力カラム順。
COLUMNS = [
    "work_id", "工事名", "設計年月", "施工都道府県", "発注事務所",
    "単価項目", "規格", "単価適用年月", "単位", "数量", "単価", "金額",
    "表種別", "摘要号", "労務判定", "ページ",
]
REVIEW_COLUMNS = COLUMNS + ["is_kept", "review理由", "検算相対誤差"]


def dedup_within_work(rows: list[dict]) -> list[dict]:
    """同一工事内の完全重複を除去。キー = (単価項目, 規格, 単位, 単価, 金額)。

    工事横断は対象外（工事 or 単価適用年月が違えば別レコードで残る）。
    """
    seen: set = set()
    kept: list[dict] = []
    for r in rows:
        key = (r.get("work_id"), r.get("単価項目"), r.get("規格"),
               r.get("単位"), r.get("単価"), r.get("金額"))
        if key in seen:
            continue
        seen.add(key)
        kept.append(r)
    return kept


def write_csv(path: Path, rows: list[dict], columns: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    # UTF-8 BOM なし。Excel 文字化け対策が必要なら encoding='utf-8-sig' に変更可。
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def main() -> None:
    ap = argparse.ArgumentParser(description="全工事統合 -> CSV")
    ap.add_argument("--out-dir", default=str(common.OUTPUT))
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    work_dirs = [p for p in sorted(common.WORK.iterdir())
                 if p.is_dir() and (p / "validated.jsonl").exists()]
    if not work_dirs:
        sys.exit("validated.jsonl がありません。先に 04_validate を実行してください。")

    db_rows: list[dict] = []
    review_rows: list[dict] = []
    for wd in work_dirs:
        rows = common.read_jsonl(wd / "validated.jsonl")
        rows = dedup_within_work(rows)
        for r in rows:
            if r.get("is_kept"):
                db_rows.append(r)
            if (not r.get("is_kept")) or r.get("review理由"):
                review_rows.append(r)

    write_csv(out_dir / "単価一覧.csv", db_rows, COLUMNS)
    write_csv(out_dir / "review_flags.csv", review_rows, REVIEW_COLUMNS)

    print(f"[consolidate] 工事{len(work_dirs)}件")
    print(f"[consolidate] DB本体 {len(db_rows)}行 -> {out_dir / '単価一覧.csv'}")
    print(f"[consolidate] レビュー {len(review_rows)}行 -> {out_dir / 'review_flags.csv'}")


if __name__ == "__main__":
    main()
