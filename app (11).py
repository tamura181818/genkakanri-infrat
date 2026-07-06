"""
見積・請求 取込アプリ（Streamlit）— 工事ごとのスプレッドシートへ自動追記

現場担当者がやること … 工事を選ぶ → 見積/請求を選ぶ → ファイルを上げる(または手入力)
   → 工種と区分を確認 → 保存。あとは台帳(スプレッドシート)が自動更新。
「API」やキーは画面に出ません。

────────────────────────────────────────────────────────
管理者（田村さん）が最初に1回だけ設定する（.streamlit/secrets.toml、GitHubには載せない）:

  [gcp_service_account]
  type = "service_account"
  project_id = "..."
  private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
  client_email = "xxxx@xxxx.iam.gserviceaccount.com"
  ...（発行されたJSONの中身をそのまま）

  GEMINI_API_KEY = "xxxx"   # 紙・写真の読み取り用（任意）

  [projects]               # 工事名 → スプレッドシートID。新しい工事は1行足す
  "中島橋補修工事" = "1AbCdEf...スプレッドシートID..."

新しい工事の手順:
  1. 「原価管理システム.xlsx」をGoogleに複製
  2. そのスプレッドシートを client_email に「編集者」で共有
  3. [projects] に 工事名 = スプレッドシートID を追加
  ※原価の閲覧可否は、各スプレッドシートを関係者だけに共有して守る（ゆるい締め）

  pip install streamlit gspread google-auth pdfplumber openpyxl pandas google-generativeai
  streamlit run app.py
"""
import re, json
import pandas as pd
import streamlit as st

st.set_page_config(page_title="見積・請求 取込", layout="wide")

KUBUN = ["材料", "機械", "労務", "経費"]
SHEET_MI, SHEET_JI, SHEET_MASTER = "見積取込", "実績取込", "工種マスタ"

KW = {
    "材料": ["材", "塗料", "ペイント", "シンナー", "プライマー", "樹脂", "鋼材", "ボルト", "セメント", "生コン", "骨材", "ネット", "シート", "資材"],
    "機械": ["機械", "重機", "リース", "賃料", "クレーン", "高所作業車", "コンプレッサ", "発電機", "バックホウ", "ブラスト", "集じん", "ポンプ", "車両", "点検車"],
    "労務": ["労務", "人件", "作業員", "世話役", "手元", "塗装工", "とび", "職長", "普通作業", "特殊作業", "交通誘導", "警備"],
}
def classify(name):
    s = str(name or "")
    for cat, words in KW.items():
        if any(w in s for w in words):
            return cat
    return "経費"

def to_num(x):
    if x is None: return 0
    s = re.sub(r"[^\d.\-]", "", str(x))
    try: return float(s) if s not in ("", "-", ".") else 0
    except: return 0

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

def _rows(rows):
    df = pd.DataFrame(rows)
    if df.empty: return pd.DataFrame(columns=["項目", "金額", "区分"])
    df["区分"] = df["項目"].map(classify)
    return df[["項目", "金額", "区分"]]

def parse_excel(file):
    xls = pd.read_excel(file, sheet_name=None, header=None)
    rows = []
    for _, df in xls.items():
        for _, r in df.iterrows():
            cells = [c for c in r.tolist() if pd.notna(c)]
            texts = [str(c) for c in cells if not isinstance(c, (int, float))]
            nums = [c for c in cells if isinstance(c, (int, float))]
            if texts and nums:
                rows.append({"項目": max(texts, key=len), "金額": max(nums)})
    return _rows(rows)

def parse_pdf(file):
    import pdfplumber
    rows, has_text, full = [], False, ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            if t: has_text = True; full += t + "\n"
            for tbl in page.extract_tables() or []:
                for r in tbl:
                    cells = [c for c in r if c not in (None, "")]
                    texts = [c for c in cells if to_num(c) == 0]
                    nums = [to_num(c) for c in cells if to_num(c) != 0]
                    if texts and nums:
                        rows.append({"項目": max(texts, key=len), "金額": max(nums)})
    if not rows and full:
        for line in full.splitlines():
            nums = [to_num(n) for n in re.findall(r"[\d,]+", line) if to_num(n) >= 100]
            name = re.sub(r"[\d,]+\s*", "", line).strip()
            if nums and len(name) >= 2:
                rows.append({"項目": name, "金額": max(nums)})
    return _rows(rows), has_text

def parse_with_gemini(file_bytes, mime, koshu):
    key = st.secrets.get("GEMINI_API_KEY", None)
    if not key: return None
    import google.generativeai as genai
    genai.configure(api_key=key)
    model = genai.GenerativeModel("gemini-2.0-flash")
    prompt = ("建設の下請け見積書/請求書です。明細を抽出し、各行に工種と区分を割り当ててください。"
              "工種リスト: " + "・".join(koshu) + "。区分: 材料/機械/労務/経費。"
              'JSONのみ返答、前置き不要。形式:[{"工種":"..","区分":"..","項目":"..","金額":0}]')
    resp = model.generate_content([{"mime_type": mime, "data": file_bytes}, prompt])
    txt = (resp.text or "").replace("```json", "").replace("```", "").strip()
    df = pd.DataFrame(json.loads(txt))
    for col, d in [("工種", ""), ("区分", "経費"), ("項目", ""), ("金額", 0)]:
        if col not in df: df[col] = d
    return df

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

up = st.file_uploader(f"{('請求' if is_seikyu else '見積')}書（Excel / PDF / 写真）を上げて自動読み取り（任意）",
                      type=["xlsx", "xls", "pdf", "png", "jpg", "jpeg"])
if "rows" not in st.session_state:
    st.session_state.rows = pd.DataFrame([{"工種": "", "区分": "材料", "項目": "", "数量": "", "単位": "", "単価": "", "金額": ""}])

if up is not None and st.button("自動読み取り"):
    ext = up.name.lower().rsplit(".", 1)[-1]
    df = None
    if ext in ("xlsx", "xls"):
        df = parse_excel(up)
    elif ext == "pdf":
        df, has_text = parse_pdf(up)
        if (not has_text or df.empty):
            ai = parse_with_gemini(up.getvalue(), "application/pdf", koshu)
            if ai is not None: df = ai
    else:
        ai = parse_with_gemini(up.getvalue(), up.type or "image/jpeg", koshu)
        if ai is not None: df = ai
    if df is None or df.empty:
        st.info("自動で読み取れませんでした。下の表に手入力してください（紙・写真は管理者のGemini設定が必要です）。")
    else:
        for col in ["工種", "数量", "単位", "単価", "金額"]:
            if col not in df: df[col] = ""
        st.session_state.rows = df.reindex(columns=["工種", "区分", "項目", "数量", "単位", "単価", "金額"]).fillna("")

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
            st.session_state.rows = pd.DataFrame([{"工種": "", "区分": "材料", "項目": "", "数量": "", "単位": "", "単価": "", "金額": ""}])
        except Exception as e:
            st.error(f"保存に失敗しました。共有と secrets 設定を確認してください。詳細: {e}")
