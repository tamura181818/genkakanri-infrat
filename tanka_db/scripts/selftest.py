#!/usr/bin/env python3
"""
selftest.py — LLM/PDF なしで決定的ロジックを検証する。

対象: 労務判定(三段構え) / 単位・和暦正規化 / 検算 / 04→05 の統合。
設計書 §1 の保持・除外例をそのままテストケース化している。

  python scripts/selftest.py
"""
from __future__ import annotations

import importlib.util
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import common  # noqa: E402


def _load(mod_file: str):
    """'04_validate.py' のように数字始まりで import 不可なモジュールを読む。"""
    spec = importlib.util.spec_from_file_location(mod_file.replace(".py", ""), HERE / mod_file)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


passed = 0
failed = 0


def check(name: str, got, want):
    global passed, failed
    if got == want:
        passed += 1
        print(f"  ok  {name}")
    else:
        failed += 1
        print(f"  NG  {name}: got={got!r} want={want!r}")


def main() -> None:
    labor = common.load_labor_exclude()
    units = common.load_unit_normalizer()
    v = _load("04_validate.py")

    print("[1] 労務判定 judge_labor（設計書 §1 の例）")
    # 保持: 労務を内包した施工単価
    check("主体足場(施工/m2)",
          common.judge_labor({"名称": "主体足場（パイプ吊足場）", "単位": "m2",
                              "単価": "3815", "金額": "3815", "labor_type": "work_unit"}, labor)[0],
          "保持")
    check("鋼材費(資材/t)",
          common.judge_labor({"名称": "鋼材費", "単位": "t", "単価": "262000",
                              "金額": "262000", "labor_type": "work_unit"}, labor)[0],
          "保持")
    # 除外: 純粋労務（辞書完全一致 + 労務単位）
    check("普通作業員(人)",
          common.judge_labor({"名称": "普通作業員", "単位": "人", "単価": "23000",
                              "金額": "23000", "labor_type": "pure_labor"}, labor)[0],
          "除外")
    check("土木一般世話役(人)",
          common.judge_labor({"名称": "土木一般世話役", "単位": "人日",
                              "単価": "28000", "金額": "28000"}, labor)[0],
          "除外")
    # 除外: 号参照（中間集計）
    check("複合単価(単-37号参照)",
          common.judge_labor({"名称": "製作加工", "単位": "t", "単価": "391600",
                              "金額": "391600", "摘要": "単-37号", "is_reference": True}, labor)[0],
          "除外")
    # 除外: 式・単価空欄の費目集計行
    check("工場製作工(式)",
          common.judge_labor({"名称": "工場製作工", "単位": "式", "単価": None,
                              "金額": None}, labor)[0],
          "除外")
    # 要レビュー: 辞書語を含むが施工単位（過剰除去防止で保持側へ）
    check("塗装工だが/m2 → 要レビュー",
          common.judge_labor({"名称": "塗装工", "単位": "m2", "単価": "1200",
                              "金額": "1200", "labor_type": "work_unit"}, labor)[0],
          "要レビュー")
    # グレー(運転費)は職種名辞書に無い → 現方針で保持
    check("移動式クレーン運転費 → 保持(現方針)",
          common.judge_labor({"名称": "移動式クレーン運転費", "単位": "日",
                              "単価": "53200", "金額": "53200", "labor_type": "work_unit"}, labor)[0],
          "保持")

    print("[2] 単位正規化 normalize_unit")
    check("㎡→m2", common.normalize_unit("㎡", units)[0], "m2")
    check("m²→m2", common.normalize_unit("m²", units)[0], "m2")
    check("供用日", common.normalize_unit("供用日", units)[0], "供用日")
    check("未知単位はknown=False", common.normalize_unit("ナニコレ", units)[1], False)

    print("[3] 和暦→西暦 normalize_era")
    check("令和8年2月", v.normalize_era("令和8年2月"), "2026-02")
    check("令和元年12月", v.normalize_era("令和元年12月"), "2019-12")
    check("2026-03", v.normalize_era("2026-03"), "2026-03")

    print("[4] 検算 checksum（許容1%）")
    check("整合(1×3815=3815)", v.checksum({"数量": "1", "単価": "3815", "金額": "3815"}, 0.01)[0], True)
    check("不整合(1×3815 vs 5000)", v.checksum({"数量": "1", "単価": "3815", "金額": "5000"}, 0.01)[0], False)
    check("金額空欄はスキップ扱い", v.checksum({"数量": "1", "単価": "3815", "金額": None}, 0.01)[0], True)

    print("[5] 04→05 統合 end-to-end（合成データ）")
    e2e_ok = run_e2e(v)
    check("end-to-end", e2e_ok, True)

    print(f"\n結果: {passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)


def run_e2e(v) -> bool:
    """合成 extracted.jsonl + meta.json を置き、04→05 を実行して CSV を検証。"""
    wid = "__selftest__"
    wdir = common.work_dir(wid)
    out_dir = common.ROOT / "work" / "__selftest_out__"
    try:
        wdir.mkdir(parents=True, exist_ok=True)
        (wdir / "meta.json").write_text(
            '{"工事名":"テスト工事","設計年月":"令和8年2月","施工都道府県":"千葉県",'
            '"発注事務所":"首都国道事務所","単価適用年月":"令和8年3月"}',
            encoding="utf-8")
        rows = [
            {"名称": "主体足場", "規格": "パイプ吊足場", "単位": "m2", "数量": "1",
             "単価": "3815", "金額": "3815", "摘要": "", "is_reference": False,
             "labor_type": "work_unit", "ページ": 50, "表種別": "1次単価表"},
            {"名称": "主体足場", "規格": "パイプ吊足場", "単位": "m2", "数量": "1",
             "単価": "3815", "金額": "3815", "摘要": "", "is_reference": False,
             "labor_type": "work_unit", "ページ": 50, "表種別": "1次単価表"},  # 完全重複
            {"名称": "普通作業員", "規格": "", "単位": "人", "数量": "1",
             "単価": "23000", "金額": "23000", "摘要": "", "is_reference": False,
             "labor_type": "pure_labor", "ページ": 51, "表種別": "1次単価表"},  # 除外
        ]
        common.write_jsonl(wdir / "extracted.jsonl", rows)

        # 04
        subprocess.run([sys.executable, str(HERE / "04_validate.py"), "--work-id", wid],
                       check=True, cwd=HERE)
        validated = common.read_jsonl(wdir / "validated.jsonl")
        kept = [r for r in validated if r["is_kept"]]
        if len(kept) != 2:  # 主体足場×2（重複除去は05）、普通作業員は除外
            print(f"    04: kept={len(kept)} (expected 2 before dedup)")
            return False
        if kept[0]["設計年月"] != "2026-02" or kept[0]["単位"] != "m2":
            print("    04: メタ/正規化不一致")
            return False

        # 05（out-dir を退避先にして本番 output/ を汚さない）
        subprocess.run([sys.executable, str(HERE / "05_consolidate.py"),
                        "--out-dir", str(out_dir)], check=True, cwd=HERE)
        db_csv = (out_dir / "単価一覧.csv").read_text(encoding="utf-8").splitlines()
        # ヘッダ + 主体足場1行（重複除去済み） = 2 行
        if len(db_csv) != 2:
            print(f"    05: 単価一覧.csv 行数={len(db_csv)} (expected 2)")
            return False
        review_csv = (out_dir / "review_flags.csv").read_text(encoding="utf-8").splitlines()
        if len(review_csv) < 2:  # 普通作業員が review に残る
            print(f"    05: review_flags.csv 行数={len(review_csv)}")
            return False
        return True
    finally:
        shutil.rmtree(wdir, ignore_errors=True)
        shutil.rmtree(out_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
