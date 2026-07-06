"""
見積・請求 取込アプリ（Streamlit）— 工事ごとのスプレッドシートへ自動追記
※ 見積書は全形式（紙・PDF・Excel）をAI（Gemini）で読み取ります。

現場担当者がやること … 工事を選ぶ → 見積/請求を選ぶ → ファイルを上げる(または手入力)
   → 工種と区分を確認 → 保存。あとは台帳(スプレッドシート)が自動更新。

管理者が最初に1回だけ設定（.streamlit/secrets.toml、GitHubには載せない）:
  [gcp_service_account] … サービスアカウントのJSONの中身
  GEMINI_API_KEY = "xxxx"                      … 見積読み取りに必須
  [projects]  "工事名" = "スプレッドシートID"   … 工事一覧

  pip install streamlit gspread google-auth pandas openpyxl google-generativeai
  streamlit run app.py
"""
import re, json
import pandas as pd
import streamlit as st

st.set_page_config(page_title="見積・請求 取込", layout="wide")

KUBUN = ["材料", "機械", "労務", "経費"]
SHEET_MI, SHEET_JI, SHEET_MASTER = "見積取込", "実績取込", "工種マスタ"
COLS = ["工種", "区分", "項目", "数量", "単位", "単価", "金額"]

def to_num(x):
    if x is None: return 0
    s = re.sub(r"[^\d.\-]", "", str(x))
    try: return float(s) if s not in ("", "-", ".") else 0
    except: return 0

# ── gspread（認証は secrets から）──
@st.cache_resource
def get_client():
    import gspread
    from google.oauth2.service_account import Credentials
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds)

def get_projects():
    return dict(st.secrets.get("projects", {}))

@st.cache_data(ttl=300)
def get_koshu(sid):
    try:
        ws = get_client().open_by_key(sid).worksheet(SHEET_MASTER)
        return [v for v in ws.col_values(1)[2:] if v]
    except Exception:
        return []

def append_rows(sid, sheet_name, rows):
    ws = get_client().open_by_key(sid).worksheet(sheet_name)
    ws.append_rows(rows, value_input_option="USER_ENTERED")

# ── 見積の読み取り（全形式をAI＝Geminiで判断）──
PROMPT = ("建設の下請け見積書です。明細行をすべて抽出してください。"
          "各行に、工種(下のリストから最も近いものを選ぶ)・区分(材料/機械/労務/経費)・"
          "項目・数量・単位・単価・金額 を付けてください。"
          "単価や数量が書かれていない『一式』の行は、金額だけ入れて数量・単価は空でよい。"
          "小計・合計・総括などの集計行は除き、実際の明細だけを返してください。"
          "工種リスト: {koshu}。"
          'JSONのみ返答、前置き・コードブロック不要。'
          '形式:[{{"工種":"","区分":"材料","項目":"","数量":0,"単位":"","単価":0,"金額":0}}]')

def _ai_call(parts, koshu):
    key = st.secrets.get("GEMINI_API_KEY", None)
    if not key:
        return None, "見積の自動読み取りには、管理者のGemini設定(GEMINI_API_KEY)が必要です。下の表に手入力してください。"
    import google.generativeai as genai
    genai.configure(api_key=key)
    model = genai.GenerativeModel("gemini-2.0-flash")
    prompt = PROMPT.format(koshu="・".join(koshu) if koshu else "（工種リストなし・推定可）")
    try:
        resp = model.generate_content(parts + [prompt])
        txt = (resp.text or "").replace("```json", "").replace("```", "").strip()
        data = json.loads(txt)
        df = pd.DataFrame(data)
        for c, d in [("工種", ""), ("区分", "材料"), ("項目", ""), ("数量", ""), ("単位", ""), ("単価", ""), ("金額", 0)]:
            if c not in df: df[c] = d
        return df[COLS], None
    except Exception as e:
        return None, f"自動読み取りに失敗しました。手入力に切り替えてください。詳細: {e}"

def parse_ai_file(file_bytes, mime, koshu):
    # PDF・画像はそのままAIへ
    return _ai_call([{"mime_type": mime, "data": file_bytes}], koshu)

def parse_ai_excel(file, koshu):
    # ExcelはAIに渡せないので、全セルをテキスト化してAIへ
    try:
        xls = pd.read_excel(file, sheet_name=None, header=None)
    except Exception as e:
        return None, f"Excelを読めませんでした: {e}"
    lines = []
    for name, df in xls.items():
        lines.append(f"[シート: {name}]")
        for _, r in df.iterrows():
            cells = [str(c) for c in r.tolist() if pd.notna(c)]
            if cells: lines.append(" | ".join(cells))
    text = "次はExcel見積書の中身です。\n" + "\n".join(lines[:400])
    return _ai_call([text], koshu)

# ────────────────────────────────────────────────────────
st.title("見積・請求 取込")

projects = get_projects()
if not projects:
    st.warning("工事が登録されていません。管理者に secrets の [projects] 設定を依頼してください。")
    st.stop()

c1, c2 = st.columns([2, 1])
koji = c1.selectbox("工事を選ぶ", list(projects.keys()))
mode = c2.radio("種別", ["見積（予算）", "請求（実績）"], horizontal=True)
is_seikyu = mode.startswith("請求")
sid = projects[koji]
koshu = get_koshu(sid)
if not koshu:
    st.info("この工事の工種マスタを読めませんでした。スプレッドシートがサービスアカウントに共有されているか確認してください。")

h1, h2, h3 = st.columns(3)
gyosha = h1.text_input("業者名")
hiduke = h2.text_input("請求日（締め日）" if is_seikyu else "見積日", placeholder="2026/07/31")
if is_seikyu:
    shiharai = h3.selectbox("支払区分", ["月締め", "出来高払い"]); ver = keiyaku = None
else:
    ver = h3.selectbox("版", ["当初", "変更", "追加"])
    keiyaku = "未契約" if ver == "追加" else "契約済"

up = st.file_uploader(f"{('請求' if is_seikyu else '見積')}書（Excel / PDF / 写真）を上げてAI読み取り（任意）",
                      type=["xlsx", "xls", "pdf", "png", "jpg", "jpeg"])
if "rows" not in st.session_state:
    st.session_state.rows = pd.DataFrame([{c: ("材料" if c == "区分" else "") for c in COLS}])

if up is not None and st.button("AIで読み取り"):
    ext = up.name.lower().rsplit(".", 1)[-1]
    with st.spinner("AIが読み取り中…"):
        if ext in ("xlsx", "xls"):
            df, err = parse_ai_excel(up, koshu)
        elif ext == "pdf":
            df, err = parse_ai_file(up.getvalue(), "application/pdf", koshu)
        else:
            df, err = parse_ai_file(up.getvalue(), up.type or "image/jpeg", koshu)
    if err:
        st.info(err)
    elif df is not None and not df.empty:
        st.session_state.rows = df.reindex(columns=COLS).fillna("")
        st.success(f"{len(df)} 行を読み取りました。工種・区分・金額を確認してください。")

st.caption("工種と区分を確認・修正してください（AIの下書きは完璧ではありません）。金額は数量×単価が空なら手入力。")
edited = st.data_editor(
    st.session_state.rows, use_container_width=True, num_rows="dynamic",
    column_config={
        "工種": st.column_config.SelectboxColumn("工種", options=koshu or [""]),
        "区分": st.column_config.SelectboxColumn("区分", options=KUBUN),
    })

def calc_amt(row):
    q, u, a = to_num(row.get("数量")), to_num(row.get("単価")), to_num(row.get("金額"))
    return q * u if (q and u) else a
edited = edited.copy()
edited["金額計"] = edited.apply(calc_amt, axis=1)

sums = {k: int(edited.loc[edited["区分"] == k, "金額計"].sum()) for k in KUBUN}
m = st.columns(5)
for i, k in enumerate(KUBUN): m[i].metric(k, f"¥{sums[k]:,}")
m[4].metric("合計", f"¥{sum(sums.values()):,}")

if st.button(f"{'請求' if is_seikyu else '見積'}を保存（{koji} のシートに追記）", type="primary"):
    valid = edited[(edited["工種"].astype(str) != "") & (edited["金額計"] > 0)]
    if valid.empty:
        st.error("工種と金額が入った行がありません。")
    else:
        try:
            if is_seikyu:
                rows = [[r["工種"], r["区分"], r["項目"], int(r["金額計"]), gyosha, hiduke, shiharai] for _, r in valid.iterrows()]
                append_rows(sid, SHEET_JI, rows)
            else:
                rows = [[r["工種"], ver, r["区分"], r["項目"], to_num(r["数量"]) or "", r["単位"],
                         to_num(r["単価"]) or "", int(r["金額計"]), gyosha, hiduke, keiyaku] for _, r in valid.iterrows()]
                append_rows(sid, SHEET_MI, rows)
            st.success(f"{koji} の「{SHEET_JI if is_seikyu else SHEET_MI}」に {len(rows)} 行を追記しました。台帳が自動更新されます。")
            st.session_state.rows = pd.DataFrame([{c: ("材料" if c == "区分" else "") for c in COLS}])
        except Exception as e:
            st.error(f"保存に失敗しました。共有と secrets 設定を確認してください。詳細: {e}")
