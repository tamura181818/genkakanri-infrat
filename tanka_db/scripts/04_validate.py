#!/usr/bin/env python3
"""
04_validate.py — 検算・労務分類確定・正規化。

設計書 §3「04_validate.py」対応。
- 検算: |数量×単価 − 金額| / 金額 が許容誤差(既定1%)超なら要レビュー。
        金額空欄の集計行はスキップ。
- 労務確定: §1 三段構え（common.judge_labor）で 保持/除外/要レビュー。
- 正規化: 単位(unit_normalize.yaml) / 全角数字→半角 / 半角カナ→全角 /
          和暦→西暦。
- メタ結合: work/{work_id}/meta.json を各行にマージし §2 データモデルへ整形。

出力: work/{work_id}/validated.jsonl   （§2 の全カラムを持つ最下層明細）
      is_kept=True の行が DB 本体、それ以外は review 対象。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import common

# ── 和暦→西暦 ────────────────────────────────────────
_ERA = {"令和": 2018, "平成": 1988, "昭和": 1925}


def normalize_era(text) -> str | None:
    """'令和8年2月' -> '2026-02'。西暦 'YYYY-MM'/'YYYY/MM' もそのまま整形。"""
    if not text:
        return None
    s = common.to_halfwidth(str(text)).strip()
    m = re.search(r"(令和|平成|昭和)\s*(\d+|元)\s*年\s*(\d+)?\s*月?", s)
    if m:
        era, y, mo = m.group(1), m.group(2), m.group(3)
        year = _ERA[era] + (1 if y == "元" else int(y))
        month = int(mo) if mo else 1
        return f"{year:04d}-{month:02d}"
    m = re.search(r"(\d{4})[-/年\.](\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    return s or None


def checksum(row: dict, tol: float) -> tuple[bool, float | None]:
    """(誤差が許容内か, 相対誤差)。金額 or 単価×数量 が取れなければ (True, None)。"""
    q = common.to_num(row.get("数量"))
    p = common.to_num(row.get("単価"))
    a = common.to_num(row.get("金額"))
    if a is None or a == 0 or q is None or p is None:
        return True, None  # 集計行等はスキップ扱い
    rel = abs(q * p - a) / abs(a)
    return rel <= tol, rel


def validate_work(work_id: str, labor_cfg: dict, unit_map: dict, tol: float) -> None:
    wdir = common.work_dir(work_id)
    rows = common.read_jsonl(wdir / "extracted.jsonl")
    if not rows:
        print(f"[validate] {work_id}: extracted.jsonl が空。先に 03_extract を実行。")
        return

    meta_path = wdir / "meta.json"
    meta = json.loads(meta_path.read_text(encoding="utf-8")) if meta_path.exists() else {}
    m_work = meta.get("工事名")
    m_design_ym = normalize_era(meta.get("設計年月"))
    m_pref = meta.get("施工都道府県")
    m_office = meta.get("発注事務所")
    m_apply_ym = normalize_era(meta.get("単価適用年月"))

    out: list[dict] = []
    n_keep = n_excl = n_review = 0

    for r in rows:
        judgment, reason = common.judge_labor(r, labor_cfg)
        ok, rel = checksum(r, tol)
        review_reasons = []
        if judgment == "要レビュー":
            review_reasons.append(f"労務判定:{reason}")
        if not ok:
            review_reasons.append(f"検算誤差{rel:.1%}>許容{tol:.0%}")

        unit_norm, unit_known = common.normalize_unit(r.get("単位"), unit_map)
        if not unit_known and unit_norm:
            review_reasons.append(f"未知単位:{unit_norm}")

        # is_kept: 保持 かつ 検算OK。除外行は DB 本体から外す（review には残す）。
        is_kept = (judgment == "保持") and ok

        rec = {
            "work_id": work_id,
            "工事名": m_work,
            "設計年月": m_design_ym,
            "施工都道府県": m_pref,
            "発注事務所": m_office,
            "単価項目": common.kana_to_fullwidth(str(r.get("名称") or "")).strip() or None,
            "規格": (common.kana_to_fullwidth(str(r.get("規格"))).strip()
                     if r.get("規格") else None),
            "単価適用年月": m_apply_ym,
            "単位": unit_norm or None,
            "数量": common.to_num(r.get("数量")),
            "単価": common.to_num(r.get("単価")),
            "金額": common.to_num(r.get("金額")),
            "表種別": r.get("表種別"),
            "摘要号": r.get("摘要号"),
            "労務判定": judgment,
            "ページ": r.get("ページ"),
            "is_kept": is_kept,
            "review理由": "; ".join(review_reasons) if review_reasons else None,
            "検算相対誤差": round(rel, 4) if rel is not None else None,
        }
        out.append(rec)
        if judgment == "除外":
            n_excl += 1
        elif judgment == "要レビュー" or not ok:
            n_review += 1
        elif is_kept:
            n_keep += 1

    common.write_jsonl(wdir / "validated.jsonl", out)
    print(f"[validate] {work_id}: 保持{n_keep} / 除外{n_excl} / 要レビュー{n_review} "
          f"-> {wdir / 'validated.jsonl'}")


def main() -> None:
    ap = argparse.ArgumentParser(description="検算・労務確定・正規化")
    ap.add_argument("--work-id", help="単一 work_id。省略時は work/ 配下すべて")
    ap.add_argument("--tol", type=float, default=0.01, help="検算許容相対誤差(既定0.01=1%%)")
    args = ap.parse_args()

    labor_cfg = common.load_labor_exclude()
    unit_map = common.load_unit_normalizer()

    if args.work_id:
        work_ids = [args.work_id]
    else:
        work_ids = [p.name for p in sorted(common.WORK.iterdir())
                    if p.is_dir() and (p / "extracted.jsonl").exists()]
    if not work_ids:
        sys.exit("処理対象がありません。先に 03_extract を実行してください。")

    for wid in work_ids:
        validate_work(wid, labor_cfg, unit_map, args.tol)


if __name__ == "__main__":
    main()
