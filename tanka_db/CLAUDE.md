# 国交省開示資料 単価データベース化システム（tanka_db）

国交省の開示設計書 PDF（1 工事 1 ファイル・約 400 工事）から**最下層の単価**を
抽出し、検索可能な CSV データベースを構築するパイプライン。設計書
`__DB__________v1.md` の実装。既存ルートの `app.py`（見積・請求取込）とは独立。

## パイプライン

```
input_pdfs/  →  01_render  →  work/{work_id}/pages/*.jpg
                02_classify (任意)      ページ種別のみ
                03_extract              種別判定＋明細抽出(ビジョンLLM) → extracted.jsonl / meta.json
                04_validate             検算・労務確定・正規化       → validated.jsonl
                05_consolidate          全工事統合                   → output/単価一覧.csv / review_flags.csv
```

## セットアップ

```bash
cd tanka_db
pip install -r requirements.txt
# 画像化に poppler-utils が必要
sudo apt-get install poppler-utils          # or: brew install poppler
export GEMINI_API_KEY=xxxx                   # 抽出に必須
export GEMINI_MODEL=gemini-2.5-flash         # 任意。flash-lite と比較検証可
```

## 実行（1 工事で検証 → 全工事へ）

```bash
# 1) PDF を input_pdfs/ に置く（Google ドライブ同期先）
cd tanka_db/scripts

# 2) 1 工事で検証
python 01_render.py --pdf ../input_pdfs/01_xxxx.pdf --dpi 200
python 03_extract.py --work-id 01_xxxx
python 04_validate.py --work-id 01_xxxx
python 05_consolidate.py

# 3) review_flags.csv を目視 → 辞書/許容誤差/カラム調整 → 全工事へ
python 01_render.py            # input_pdfs/ の全 PDF
python 03_extract.py
python 04_validate.py
python 05_consolidate.py
```

## 抽出スコープ（設計書 §1）

最下層の単価のみを保持する。**号参照（単-○号／内-○号）を持つ行**と
**純粋労務（職種の賃金）** は除外。判定は三段構え:

1. 号参照フィルタ: 摘要に `単-`/`内-`、または単価空欄で単位「式」→ 除外（中間集計）。
2. 除外辞書 × 単位併用: 名称が `config/labor_exclude.yaml` に一致し労務単位なら除外。
   施工単位なら保持側へ倒して要レビュー。
3. LLM 分類（`labor_type`）: `pure_labor`/`work_unit`/`unknown` を材料に確定。
   確信が持てない行は「要レビュー」。

実装は `scripts/common.py: judge_labor()`。

## 設定ファイル（コード変更なしで調整）

- `config/labor_exclude.yaml` … 純粋労務の除外辞書。拾いすぎ/取りこぼしで語を増減。
- `config/unit_normalize.yaml` … 単位の表記ゆれ辞書（m²/㎡→m2 等）。

## 出力（設計書 §2 データモデル）

`output/単価一覧.csv`（UTF-8 BOM なし）のカラム:
`work_id, 工事名, 設計年月, 施工都道府県, 発注事務所, 単価項目, 規格,
単価適用年月, 単位, 数量, 単価, 金額, 表種別, 摘要号, 労務判定, ページ`

要レビュー・検算超過・除外行は `output/review_flags.csv` に分離。

重複方針: 同一工事内の完全重複は除去。工事・単価適用年月が異なる同一品目は
別レコードとして残す（時系列・事務所間比較のため）。

## 未確定・要決定（設計書 §6）

運転費の除外可否 / 単価適用年月カラム / 抽出モデル(Flash vs Flash-Lite) /
DPI(150 vs 200) / 重複方針 / 出力先(CSV or gspread) / 検索 UI(GAS or Streamlit) /
様式ゆれ。1 工事検証後に `review_flags.csv` を見ながら田村さんと確定する。

## 自己テスト

`python scripts/selftest.py` で決定的ロジック（労務判定・正規化・検算・統合）を
LLM/PDF なしで検証できる。
