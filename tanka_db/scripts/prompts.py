"""
ビジョン LLM(Gemini 想定) 用のプロンプトと structured-output スキーマ。

設計書 §3「02_classify / 03_extract」対応。
02_classify と 03_extract の統合方針（1 リクエストで種別判定＋明細抽出）を
既定とし、ページ種別と明細行を同時に返させる。
"""

PAGE_TYPES = ["鏡", "設計内訳書", "内訳書", "単価表", "その他"]

# ── 明細ページ: 種別判定 + 明細抽出（統合） ─────────────
EXTRACT_PROMPT = """あなたは国土交通省の開示設計書PDFを読む専門家です。
渡された画像は工事設計書の1ページ（A4・正立済み）です。次を行ってください。

1) このページの種別を次から1つ選ぶ: 鏡 / 設計内訳書 / 内訳書 / 単価表 / その他
   - 鏡: 1ページ目のメタ情報（工事名・工事地名・発注年月・事務所名・設計年月・
         施工県・単価適用年月・路線 等）。表の明細は無い。
   - 設計内訳書: 列が〔工事区分・工種・種別・細別｜規格｜単位｜数量｜単価｜金額｜
         数量増減｜金額増減｜摘要〕の内訳表。
   - 内訳書: 「一式当たり内訳書（内-○号）」。
   - 単価表: 「1次・2次単価表（単-○号）」。
   - その他: 表紙・目次・図面など。

2) 表の明細行をすべて抽出する（ヘッダ行・区切り線は除く）。各行:
   { "名称", "規格", "単位", "数量", "単価", "金額", "摘要",
     "is_reference", "labor_type" }
   ルール:
   - 「計 / 合計 / 小計 / 総括」や、単価が空欄で単位が「式」の集計行は
     明細に含めない（除外）。
   - 摘要に「単-」または「内-」を含む行は is_reference=true。それ以外 false。
   - labor_type: 職種の賃金・直接労務費そのものなら "pure_labor"、
     労務を内包した作業/施工単価なら "work_unit"、判別不能なら "unknown"。
     （例: 普通作業員=pure_labor、主体足場=work_unit、鋼材費=work_unit）
   - 半角カナはそのまま出力（正規化は後段で行う）。読めない値は null。
   - 表種別が「設計内訳書」「その他」で最下層明細が無い場合は rows を空配列に。

出力は指定スキーマの JSON のみ。前置き・コードブロックは不要。
"""

# ── 鏡ページ: メタ情報抽出 ──────────────────────────────
KAGAMI_PROMPT = """渡された画像は国土交通省 開示設計書の「鏡」（1ページ目）です。
次のメタ情報を読み取り、JSON で返してください。読めない項目は null。

{ "工事名", "工事地名", "発注年月", "発注事務所", "設計年月",
  "施工都道府県", "単価適用年月", "路線" }

- 「施工県」は「施工都道府県」に対応。
- 年月は原文のまま（和暦でも可。西暦正規化は後段で行う）。
出力は JSON のみ。前置き・コードブロック不要。
"""

# ── structured output スキーマ（google-generativeai response_schema） ──
# OpenAPI 風スキーマ。SDK バージョン差異に備え、dict で表現する。
LINE_ITEM_SCHEMA = {
    "type": "object",
    "properties": {
        "名称": {"type": "string", "nullable": True},
        "規格": {"type": "string", "nullable": True},
        "単位": {"type": "string", "nullable": True},
        "数量": {"type": "string", "nullable": True},
        "単価": {"type": "string", "nullable": True},
        "金額": {"type": "string", "nullable": True},
        "摘要": {"type": "string", "nullable": True},
        "is_reference": {"type": "boolean"},
        "labor_type": {
            "type": "string",
            "enum": ["pure_labor", "work_unit", "unknown"],
        },
    },
    "required": ["名称", "is_reference", "labor_type"],
}

EXTRACT_SCHEMA = {
    "type": "object",
    "properties": {
        "page_type": {"type": "string", "enum": PAGE_TYPES},
        "摘要号": {"type": "string", "nullable": True},  # このページが 単-○号/内-○号 なら
        "rows": {"type": "array", "items": LINE_ITEM_SCHEMA},
    },
    "required": ["page_type", "rows"],
}

KAGAMI_SCHEMA = {
    "type": "object",
    "properties": {
        "工事名": {"type": "string", "nullable": True},
        "工事地名": {"type": "string", "nullable": True},
        "発注年月": {"type": "string", "nullable": True},
        "発注事務所": {"type": "string", "nullable": True},
        "設計年月": {"type": "string", "nullable": True},
        "施工都道府県": {"type": "string", "nullable": True},
        "単価適用年月": {"type": "string", "nullable": True},
        "路線": {"type": "string", "nullable": True},
    },
}
