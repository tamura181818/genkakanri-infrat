"""
見積・請求 取込アプリ + 原価管理ダッシュボード（Streamlit）
※ 見積書は全形式（紙・PDF・Excel）をAI（Gemini）で読み取り、工事ごとのスプレッドシートへ自動追記します。
※ 「原価管理」タブでは、溜まった見積(予算)・実績(原価)から 工種別・区分別の予実対比を表示します。

現場担当者がやること … 工事を選ぶ → 見積/請求を選ぶ → ファイルを上げる(または手入力)
   → 工種と区分を確認 → 保存。あとは台帳(スプレッドシート)が自動更新。
管理者・所長 … 「原価管理」タブで 予算残・消化率・超過・粗利 を確認。

管理者が最初に1回だけ設定（.streamlit/secrets.toml、GitHubには載せない）:
  [gcp_service_account] … サービスアカウントのJSONの中身
  GEMINI_API_KEY = "xxxx"                      … 見積読み取りに必須
  [projects]  "工事名" = "スプレッドシートID"   … 工事一覧
  [uketori]   "工事名" = 12345678              … （任意）請負金額(税抜)。入れると粗利が出ます

  pip install streamlit gspread google-auth pandas openpyxl google-generativeai
  streamlit run app.py
"""
import re, json
import pandas as pd
import streamlit as st

st.set_page_config(page_title="見積・請求 取込 / 原価管理", layout="wide")

KUBUN = ["材料", "機械", "労務", "経費"]
SHEET_MI, SHEET_JI, SHEET_MASTER = "見積取込", "実績取込", "工種マスタ"
COLS = ["工種", "区分", "項目", "数量", "単位", "単価", "金額"]

def to_num(x):
    if x is None: return 0
    s = re.sub(r"[^\d.\-]", "", str(x))
    try: return float(s) if s not in ("", "-", ".") else 0
    except: return 0

def default_rows():
    return pd.DataFrame([{c: ("材料" if c == "区分" else "") for c in COLS}])

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

def get_uketori(koji):
    """請負金額(受注・税抜)。secrets の [uketori] にあれば返す。なければ 0。"""
    return to_num(dict(st.secrets.get("uketori", {})).get(koji, 0))

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

# ── シート読み込み（原価管理ダッシュボード用）──
@st.cache_data(ttl=120)
def read_records(sid, sheet_name):
    """シートを DataFrame で返す。ヘッダ行（工種・区分を含む行）を自動検出する。"""
    try:
        ws = get_client().open_by_key(sid).worksheet(sheet_name)
        values = ws.get_all_values()
    except Exception as e:
        return None, str(e)
    if not values:
        return pd.DataFrame(), None
    header_idx = 0
    for i, row in enumerate(values[:6]):
        if "工種" in row and "区分" in row:
            header_idx = i
            break
    header = values[header_idx]
    data = values[header_idx + 1:]
    # 列名の重複・空を安全化
    seen, cols = {}, []
    for j, h in enumerate(header):
        name = h.strip() or f"col{j}"
        if name in seen:
            seen[name] += 1; name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        cols.append(name)
    df = pd.DataFrame(data, columns=cols)
    return df, None

def pivot_amount(df, amount_col="金額"):
    """工種×区分ごとに金額を合算した Series を返す。"""
    if df is None or df.empty:
        return pd.Series(dtype=float)
    if not ({"工種", "区分", amount_col} <= set(df.columns)):
        return pd.Series(dtype=float)
    d = df.copy()
    d[amount_col] = d[amount_col].apply(to_num)
    d = d[d["工種"].astype(str).str.strip() != ""]
    if d.empty:
        return pd.Series(dtype=float)
    return d.groupby(["工種", "区分"])[amount_col].sum()

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
    total = len(lines)
    text = "次はExcel見積書の中身です。\n" + "\n".join(lines[:800])
    warn = "（Excelが大きいため一部のみAIに渡しました。金額の抜けにご注意ください）" if total > 800 else None
    df, err = _ai_call([text], koshu)
    if warn and not err:
        st.warning(warn)
    return df, err

# ────────────────────────────────────────────────────────
# 取込タブ
# ────────────────────────────────────────────────────────
def render_import(koji, sid, koshu):
    mode = st.radio("種別", ["見積（予算）", "請求（実績）"], horizontal=True)
    is_seikyu = mode.startswith("請求")

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

    # 明細合計 と 手入力の伝票総額 の突合（拾い漏れ・二重計上の早期発見）
    total = sum(sums.values())
    denpyo = st.number_input("伝票の総額（任意・突合用）", min_value=0, value=0, step=1000,
                             help="見積書・請求書に書かれた合計額を入れると、明細合計とのズレを表示します。")
    if denpyo:
        diff = total - int(denpyo)
        if diff == 0:
            st.success("明細合計と伝票総額が一致しています。")
        else:
            st.warning(f"明細合計と伝票総額に {abs(diff):,} 円のズレがあります（明細 {total:,} / 伝票 {int(denpyo):,}）。値引き・一式・拾い漏れをご確認ください。")

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
                read_records.clear()  # ダッシュボードのキャッシュを更新
                st.session_state.rows = default_rows()
            except Exception as e:
                st.error(f"保存に失敗しました。共有と secrets 設定を確認してください。詳細: {e}")

# ────────────────────────────────────────────────────────
# 原価管理タブ（予実対比ダッシュボード）
# ────────────────────────────────────────────────────────
def render_dashboard(koji, sid):
    top = st.columns([3, 1])
    top[0].subheader(f"📊 {koji} の原価管理（予実対比）")
    if top[1].button("🔄 最新に更新"):
        read_records.clear()

    bud_raw, e1 = read_records(sid, SHEET_MI)
    act_raw, e2 = read_records(sid, SHEET_JI)
    if e1: st.warning(f"「{SHEET_MI}」を読めませんでした: {e1}")
    if e2: st.warning(f"「{SHEET_JI}」を読めませんでした: {e2}")
    bud = bud_raw if bud_raw is not None else pd.DataFrame()
    act = act_raw if act_raw is not None else pd.DataFrame()

    # 実行予算の対象（見積の版）を選択：当初のみ / 当初+変更 / すべて など
    if not bud.empty and "版" in bud.columns:
        present = [v for v in ["当初", "変更", "追加"] if v in set(bud["版"].astype(str))]
        sel = st.multiselect("実行予算に含める見積の版", present, default=present,
                             help="通常は当初＋変更＋追加の合計を実行予算とみなします。当初だけに絞ることもできます。")
        if sel:
            bud = bud[bud["版"].astype(str).isin(sel)]

    bpiv = pivot_amount(bud)
    apiv = pivot_amount(act)
    m = pd.DataFrame({"実行予算": bpiv, "実績原価": apiv}).fillna(0)
    if m.empty:
        st.info("まだデータがありません。「取込」タブで見積・請求を保存してください。")
        return
    m = m.reset_index().rename(columns={"level_0": "工種", "level_1": "区分"})
    if "工種" not in m.columns:  # index名が付かない環境向けの保険
        m.columns = ["工種", "区分", "実行予算", "実績原価"]
    m["差異(残予算)"] = m["実行予算"] - m["実績原価"]
    m["消化率"] = m.apply(lambda r: (r["実績原価"] / r["実行予算"]) if r["実行予算"] else (1.0 if r["実績原価"] else 0.0), axis=1)

    tot_b, tot_a = m["実行予算"].sum(), m["実績原価"].sum()

    # ── KPIタイル ──
    k = st.columns(4)
    k[0].metric("実行予算", f"¥{int(tot_b):,}")
    k[1].metric("実績原価", f"¥{int(tot_a):,}")
    k[2].metric("残予算", f"¥{int(tot_b - tot_a):,}")
    k[3].metric("消化率", f"{(tot_a / tot_b * 100 if tot_b else 0):.1f}%")

    # ── 粗利（請負金額があれば）──
    default_uke = int(get_uketori(koji))
    uketori = st.number_input("請負金額（受注・税抜）を入れると粗利が出ます", min_value=0, value=default_uke, step=100000)
    if uketori:
        g = st.columns(3)
        g[0].metric("粗利（実績ベース）", f"¥{int(uketori - tot_a):,}", f"粗利率 {((uketori - tot_a) / uketori * 100):.1f}%")
        g[1].metric("粗利（予算ベース）", f"¥{int(uketori - tot_b):,}", f"粗利率 {((uketori - tot_b) / uketori * 100):.1f}%")
        g[2].metric("原価率（実績）", f"{(tot_a / uketori * 100):.1f}%")

    # ── 超過アラート ──
    over = m[m["消化率"] > 1.0].sort_values("消化率", ascending=False)
    if not over.empty:
        st.error(f"⚠ 予算超過が {len(over)} 件あります（工種・区分別）")
        alert = over[["工種", "区分", "実行予算", "実績原価", "差異(残予算)"]].copy()
        for c in ["実行予算", "実績原価", "差異(残予算)"]:
            alert[c] = alert[c].map(lambda v: f"¥{int(v):,}")
        st.dataframe(alert, use_container_width=True, hide_index=True)

    # ── 区分別サマリ ──
    st.markdown("#### 区分別サマリ")
    byk = m.groupby("区分")[["実行予算", "実績原価"]].sum().reindex(KUBUN).fillna(0)
    cols = st.columns(len(KUBUN))
    for i, kbn in enumerate(KUBUN):
        b, a = int(byk.loc[kbn, "実行予算"]), int(byk.loc[kbn, "実績原価"])
        rate = f"{(a / b * 100):.0f}%" if b else "-"
        cols[i].metric(kbn, f"¥{a:,} / ¥{b:,}", f"消化 {rate}", delta_color="off")

    # ── 工種別 予実テーブル ──
    st.markdown("#### 工種・区分別 予実対比")
    disp = m.sort_values(["工種", "区分"]).copy()
    disp["状態"] = disp["消化率"].map(lambda x: "⚠ 超過" if x > 1.0 else ("● 要注意" if x >= 0.9 else ""))
    disp["消化率"] = (disp["消化率"] * 100).round(1).map(lambda v: f"{v}%")
    for c in ["実行予算", "実績原価", "差異(残予算)"]:
        disp[c] = disp[c].map(lambda v: f"¥{int(v):,}")
    st.dataframe(disp[["工種", "区分", "実行予算", "実績原価", "差異(残予算)", "消化率", "状態"]],
                 use_container_width=True, hide_index=True)

    # ── 工種別チャート ──
    st.markdown("#### 工種別 予算 vs 実績")
    chart = m.groupby("工種")[["実行予算", "実績原価"]].sum()
    st.bar_chart(chart)

# ────────────────────────────────────────────────────────
st.title("見積・請求 取込 / 原価管理")

projects = get_projects()
if not projects:
    st.warning("工事が登録されていません。管理者に secrets の [projects] 設定を依頼してください。")
    st.stop()

koji = st.selectbox("工事を選ぶ", list(projects.keys()))
sid = projects[koji]

# 工事を切り替えたら入力途中データをクリア（別工事への誤保存を防止）
if st.session_state.get("cur_koji") != koji:
    st.session_state.cur_koji = koji
    st.session_state.rows = default_rows()
if "rows" not in st.session_state:
    st.session_state.rows = default_rows()

koshu = get_koshu(sid)
if not koshu:
    st.info("この工事の工種マスタを読めませんでした。スプレッドシートがサービスアカウントに共有されているか確認してください。")

tab_import, tab_dash = st.tabs(["📥 取込", "📊 原価管理"])
with tab_import:
    render_import(koji, sid, koshu)
with tab_dash:
    render_dashboard(koji, sid)
