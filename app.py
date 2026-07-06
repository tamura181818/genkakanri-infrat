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
from datetime import date
import pandas as pd
import streamlit as st

st.set_page_config(page_title="見積・請求 取込 / 原価管理", layout="wide")

KUBUN = ["材料", "機械", "労務", "経費"]
SHEET_MI, SHEET_JI, SHEET_MASTER = "見積取込", "実績取込", "工種マスタ"
SHEET_YOSAN = "実行予算"
COLS = ["工種", "区分", "項目", "数量", "単位", "単価", "金額"]
DEFAULT_TAX = 10  # 税率(%) 既定値

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

def get_or_create_ws(sid, name):
    sh = get_client().open_by_key(sid)
    try:
        return sh.worksheet(name)
    except Exception:
        return sh.add_worksheet(title=name, rows=200, cols=12)

def replace_sheet(sid, name, header, rows):
    """シート全体をヘッダ＋rowsで置き換える（実行予算の凍結に使用）。"""
    ws = get_or_create_ws(sid, name)
    ws.clear()
    ws.update([header] + rows, value_input_option="USER_ENTERED")

def count_existing(sid, sheet_name, gyosha, date_col, date_val):
    """同一業者・同一日付の既存行数を返す（重複取込の検知）。"""
    gyosha, date_val = str(gyosha).strip(), str(date_val).strip()
    if not gyosha or not date_val:
        return 0
    df, err = read_records(sid, sheet_name)
    if err or df is None or df.empty:
        return 0
    if "業者名" not in df.columns or date_col not in df.columns:
        return 0
    mask = (df["業者名"].astype(str).str.strip() == gyosha) & (df[date_col].astype(str).str.strip() == date_val)
    return int(mask.sum())

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

    # 税区分：内部は必ず税抜で保存し、予実を同じ土俵で比較する
    t1, t2 = st.columns([2, 1])
    zeikei = t1.radio("入力金額の税区分", ["税抜", "税込"], horizontal=True,
                      help="請求書・見積書の金額が税込の場合は『税込』を選ぶと、税抜に換算して保存します。")
    zeiritsu = t2.number_input("税率(%)", min_value=0, max_value=20, value=DEFAULT_TAX) if zeikei == "税込" else DEFAULT_TAX

    def to_net(v):
        n = to_num(v)
        return n / (1 + zeiritsu / 100) if zeikei == "税込" else n

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
    edited["金額計"] = edited.apply(calc_amt, axis=1)          # 入力ベース（税抜/税込そのまま）
    edited["税抜金額"] = edited["金額計"].apply(to_net)         # 保存・集計は税抜に統一

    sums = {k: int(edited.loc[edited["区分"] == k, "税抜金額"].sum()) for k in KUBUN}
    m = st.columns(5)
    for i, k in enumerate(KUBUN): m[i].metric(k, f"¥{sums[k]:,}")
    m[4].metric("合計（税抜）", f"¥{sum(sums.values()):,}")
    if zeikei == "税込":
        st.caption(f"税込入力を税率{int(zeiritsu)}%で税抜換算して集計・保存します。")

    # 明細合計 と 手入力の伝票総額 の突合（拾い漏れ・二重計上の早期発見）※入力ベースで比較
    total_raw = int(edited["金額計"].sum())
    denpyo = st.number_input(f"伝票の総額（任意・突合用／{zeikei}）", min_value=0, value=0, step=1000,
                             help="見積書・請求書に書かれた合計額（入力金額と同じ税区分）を入れると、明細合計とのズレを表示します。")
    if denpyo:
        diff = total_raw - int(denpyo)
        if diff == 0:
            st.success("明細合計と伝票総額が一致しています。")
        else:
            st.warning(f"明細合計と伝票総額に {abs(diff):,} 円のズレがあります（明細 {total_raw:,} / 伝票 {int(denpyo):,}）。値引き・一式・拾い漏れをご確認ください。")

    # 重複取込チェック：同一業者・同一日付が既にあれば警告し、承知の上でのみ保存を許可
    target_sheet = SHEET_JI if is_seikyu else SHEET_MI
    date_col = "請求日" if is_seikyu else "見積日"
    dupe_n = count_existing(sid, target_sheet, gyosha, date_col, hiduke)
    allow_dupe = True
    if dupe_n:
        st.warning(f"⚠ 同じ業者「{gyosha}」・{date_col}「{hiduke}」のデータが既に {dupe_n} 行あります。二重計上の可能性があります。")
        allow_dupe = st.checkbox("重複を承知の上で保存する")

    if st.button(f"{'請求' if is_seikyu else '見積'}を保存（{koji} のシートに追記）", type="primary"):
        if dupe_n and not allow_dupe:
            st.error("重複の可能性があります。内容を確認し、問題なければ『重複を承知の上で保存する』にチェックしてください。")
        else:
            valid = edited[(edited["工種"].astype(str) != "") & (edited["税抜金額"] > 0)]
            if valid.empty:
                st.error("工種と金額が入った行がありません。")
            else:
                try:
                    if is_seikyu:
                        rows = [[r["工種"], r["区分"], r["項目"], int(round(r["税抜金額"])), gyosha, hiduke, shiharai] for _, r in valid.iterrows()]
                        append_rows(sid, SHEET_JI, rows)
                    else:
                        rows = [[r["工種"], ver, r["区分"], r["項目"], to_num(r["数量"]) or "", r["単位"],
                                 int(round(to_net(r["単価"]))) or "", int(round(r["税抜金額"])), gyosha, hiduke, keiyaku] for _, r in valid.iterrows()]
                        append_rows(sid, SHEET_MI, rows)
                    st.success(f"{koji} の「{target_sheet}」に {len(rows)} 行を税抜で追記しました。台帳が自動更新されます。")
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
    yosan_raw, _ = read_records(sid, SHEET_YOSAN)
    if e1: st.warning(f"「{SHEET_MI}」を読めませんでした: {e1}")
    if e2: st.warning(f"「{SHEET_JI}」を読めませんでした: {e2}")
    bud = bud_raw if bud_raw is not None else pd.DataFrame()
    act = act_raw if act_raw is not None else pd.DataFrame()

    # 基準＝確定済みの実行予算。無ければ見積を予算とみなす（フォールバック）
    use_yosan = (yosan_raw is not None and not yosan_raw.empty and {"工種", "区分", "予算金額"} <= set(yosan_raw.columns))
    if use_yosan:
        bpiv = pivot_amount(yosan_raw, "予算金額")
        st.caption("基準＝**実行予算（確定）**。実行予算 vs 実績原価で対比しています。")
    else:
        # 実行予算 未確定：見積を予算とみなす。版で範囲を選択可能。
        if not bud.empty and "版" in bud.columns:
            present = [v for v in ["当初", "変更", "追加"] if v in set(bud["版"].astype(str))]
            sel = st.multiselect("見積の版（実行予算 未確定のため見積を予算とみなします）", present, default=present,
                                 help="『実行予算』タブで確定すると、この見積フォールバックではなく確定予算が基準になります。")
            if sel:
                bud = bud[bud["版"].astype(str).isin(sel)]
        bpiv = pivot_amount(bud)
        st.info("実行予算が未確定のため、**見積を予算とみなして**表示しています。『🎯 実行予算』タブで確定すると精度が上がります。")

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
# 実行予算タブ（見積から作成 → 調整 → 確定/凍結）
# ────────────────────────────────────────────────────────
def render_budget(koji, sid):
    st.subheader(f"🎯 {koji} の実行予算")
    st.caption("見積を集計して実行予算のたたき台を作り、金額を調整して『確定（凍結）』すると、以後の原価管理はこの実行予算を基準に予実対比します。金額はすべて税抜です。")

    bud_raw, e1 = read_records(sid, SHEET_MI)
    if e1: st.warning(f"「{SHEET_MI}」を読めませんでした: {e1}")
    bud = bud_raw if bud_raw is not None else pd.DataFrame()
    est = pivot_amount(bud)  # 見積 by 工種×区分（税抜）

    yosan_raw, _ = read_records(sid, SHEET_YOSAN)
    frozen = {}
    if yosan_raw is not None and not yosan_raw.empty and {"工種", "区分", "予算金額"} <= set(yosan_raw.columns):
        for _, r in yosan_raw.iterrows():
            frozen[(str(r["工種"]).strip(), str(r["区分"]).strip())] = to_num(r["予算金額"])
        conf = yosan_raw["確定日"].dropna().astype(str) if "確定日" in yosan_raw.columns else pd.Series([], dtype=str)
        st.success(f"実行予算は確定済みです（{len(frozen)}件{'・確定日 ' + conf.iloc[-1] if not conf.empty else ''}）。金額を直して再確定できます。")

    keys = set(est.index) | set(frozen.keys())
    if not keys:
        st.info("見積データがありません。先に『📥 取込』タブで見積を保存してください。")
        return

    base = []
    for (k, ku) in sorted(keys):
        mitsumori = int(est.get((k, ku), 0))
        yosan = int(frozen.get((k, ku), mitsumori))
        base.append({"工種": k, "区分": ku, "見積合計": mitsumori, "予算金額": yosan})
    base_df = pd.DataFrame(base)

    st.markdown("#### 工種・区分別の実行予算（予算金額を調整できます・行の追加も可）")
    edited = st.data_editor(
        base_df, use_container_width=True, num_rows="dynamic",
        column_config={
            "工種": st.column_config.SelectboxColumn("工種", options=get_koshu(sid) or [""]),
            "区分": st.column_config.SelectboxColumn("区分", options=KUBUN),
            "見積合計": st.column_config.NumberColumn("見積合計", disabled=True, format="¥%d"),
            "予算金額": st.column_config.NumberColumn("予算金額", format="¥%d"),
        })

    yobi = st.number_input("予備費（工種『予備費』・区分『経費』として計上）", min_value=0, value=0, step=100000)
    tot_est = int(pd.Series([to_num(v) for v in edited["見積合計"]]).sum())
    tot_yosan = int(pd.Series([to_num(v) for v in edited["予算金額"]]).sum()) + int(yobi)
    c = st.columns(3)
    c[0].metric("見積合計（税抜）", f"¥{tot_est:,}")
    c[1].metric("実行予算 計（税抜）", f"¥{tot_yosan:,}")
    c[2].metric("見積との差", f"¥{tot_yosan - tot_est:,}")

    st.warning("『確定』すると実行予算シート全体を上書きします（既存の実行予算は置き換わります）。")
    if st.button("実行予算を確定（凍結）", type="primary"):
        rows, today = [], date.today().isoformat()
        for _, r in edited.iterrows():
            k, ku, amt = str(r["工種"]).strip(), str(r["区分"]).strip(), int(round(to_num(r["予算金額"])))
            if k and amt > 0:
                rows.append([k, ku or "経費", amt, today])
        if yobi > 0:
            rows.append(["予備費", "経費", int(yobi), today])
        if not rows:
            st.error("予算金額の入った行がありません。")
        else:
            try:
                replace_sheet(sid, SHEET_YOSAN, ["工種", "区分", "予算金額", "確定日"], rows)
                read_records.clear()
                st.success(f"実行予算を確定しました（{len(rows)}件）。『📊 原価管理』タブがこの実行予算を基準に対比します。")
            except Exception as e:
                st.error(f"確定に失敗しました。共有と secrets 設定を確認してください。詳細: {e}")

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

tab_import, tab_budget, tab_dash = st.tabs(["📥 取込", "🎯 実行予算", "📊 原価管理"])
with tab_import:
    render_import(koji, sid, koshu)
with tab_budget:
    render_budget(koji, sid)
with tab_dash:
    render_dashboard(koji, sid)
