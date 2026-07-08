"""
見積・請求 取込アプリ（Streamlit）— 工事ごとのスプレッドシートへ安全に追記する「入力ツール」
※ 見積書は全形式（紙・PDF・Excel）をAI（Gemini）で読み取ります。
※ 予実・粗利・支払予定などの集計は、スプレッドシート側の台帳（総合管理表・工種別内訳・
   支払予定 など）が数式で自動計算します。このアプリはその入力（見積取込・実績取込への
   追記）に専念します。

現場担当者がやること … 工事を選ぶ → 見積/請求を選ぶ → ファイルを上げる(または手入力)
   → 工種と区分を確認 → 保存。あとは台帳(スプレッドシート)が自動更新。

安全機能 … 税抜への統一・明細と伝票総額の突合・同一業者/日付の重複取込チェック・
           誤読データの修正/取消。

管理者が最初に1回だけ設定（.streamlit/secrets.toml、GitHubには載せない）:
  [gcp_service_account] … サービスアカウントのJSONの中身
  GEMINI_API_KEY = "xxxx"                      … 見積読み取りに必須
  [projects]  "工事名" = "スプレッドシートID"   … 工事一覧

  pip install streamlit gspread google-auth pandas openpyxl google-generativeai
  streamlit run app.py

※ スプレッドシートの実列（この構造に合わせています）:
  見積取込: 工種 / 版 / 区分 / 項目 / 数量 / 単位 / 単価 / 金額 / 業者 / 日付 / 契約状況
  実績取込: 工種 / 区分 / 項目 / 金額 / 業者 / 請求日 / 支払区分
  工種マスタ: 1列目に工種（3行目以降）
  （いずれも1行目はタイトル、2行目が見出し）
"""
import re, json
import pandas as pd
import streamlit as st

st.set_page_config(page_title="見積・請求 取込", layout="wide")

KUBUN = ["材料", "機械", "労務", "経費"]
SHEET_MI, SHEET_JI, SHEET_MASTER = "見積取込", "実績取込", "工種マスタ"
COLS = ["工種", "区分", "項目", "数量", "単位", "単価", "金額"]
DEFAULT_TAX = 10  # 税率(%) 既定値
BAK_PREFIX = "_bak_"  # 上書き前スナップショットの退避タブ接頭辞

def to_num(x):
    if x is None: return 0
    s = re.sub(r"[^\d.\-]", "", str(x))
    try: return float(s) if s not in ("", "-", ".") else 0
    except: return 0

def norm_date(x):
    """日付表記のゆれ（2026/07/31 と 2026-07-31 00:00:00 等）を YYYYMMDD に正規化。"""
    d = re.sub(r"\D", "", str(x))
    return d[:8]

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

def overwrite_ws(sid, name, matrix):
    """シート全体を matrix（タイトル・ヘッダ含む行の配列）で置き換える（修正・取消に使用）。"""
    ws = get_or_create_ws(sid, name)
    ws.clear()
    if matrix:
        ws.update(matrix, value_input_option="USER_ENTERED")

def snapshot_sheet(sid, sheet_name):
    """上書きの直前に、対象シートの現在値を退避タブ（_bak_<シート名>）へ1世代保存する。"""
    vals, _, err = read_sheet_raw(sid, sheet_name)
    if err or not vals:
        return False
    overwrite_ws(sid, BAK_PREFIX + sheet_name, vals)
    return True

def uniquify(header):
    seen, cols = {}, []
    for j, h in enumerate(header):
        name = str(h).strip() or f"col{j}"
        if name in seen:
            seen[name] += 1; name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        cols.append(name)
    return cols

@st.cache_data(ttl=120)
def read_sheet_raw(sid, sheet_name):
    """シートの全セルとヘッダ行位置を返す。(values, header_idx, err)。"""
    try:
        ws = get_client().open_by_key(sid).worksheet(sheet_name)
        values = ws.get_all_values()
    except Exception as e:
        return None, 0, str(e)
    if not values:
        return [], 0, None
    header_idx = 0
    for i, row in enumerate(values[:6]):
        if "工種" in row:
            header_idx = i
            break
    return values, header_idx, None

def read_records(sid, sheet_name):
    """シートを DataFrame で返す（ヘッダ行を自動検出）。"""
    values, header_idx, err = read_sheet_raw(sid, sheet_name)
    if err:
        return None, err
    if not values:
        return pd.DataFrame(), None
    header = uniquify(values[header_idx])
    width = len(header)
    data = [(row + [""] * width)[:width] for row in values[header_idx + 1:]]
    return pd.DataFrame(data, columns=header), None

def clear_data_cache():
    read_sheet_raw.clear()

def count_existing(sid, sheet_name, gyosha, date_col, date_val, gyosha_col="業者"):
    """同一業者・同一日付の既存行数を返す（重複取込の検知）。日付は表記ゆれを吸収して比較。"""
    gyosha, key = str(gyosha).strip(), norm_date(date_val)
    if not gyosha or not key:
        return 0
    df, err = read_records(sid, sheet_name)
    if err or df is None or df.empty:
        return 0
    if gyosha_col not in df.columns or date_col not in df.columns:
        return 0
    mask = (df[gyosha_col].astype(str).str.strip() == gyosha) & (df[date_col].apply(norm_date) == key)
    return int(mask.sum())

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
    return _ai_call([{"mime_type": mime, "data": file_bytes}], koshu)

def parse_ai_excel(file, koshu):
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
# 取込タブ（見積取込 / 実績取込 への追記）
# ────────────────────────────────────────────────────────
def render_import(koji, sid, koshu):
    mode = st.radio("種別", ["見積（予算）", "請求（実績）"], horizontal=True)
    is_seikyu = mode.startswith("請求")

    h1, h2, h3 = st.columns(3)
    gyosha = h1.text_input("業者")
    hiduke = h2.text_input("請求日（締め日）" if is_seikyu else "日付（見積日）", placeholder="2026/07/31")
    if is_seikyu:
        shiharai = h3.selectbox("支払区分", ["月締め", "出来高払い"]); ver = keiyaku = None
    else:
        ver = h3.selectbox("版", ["当初", "変更", "追加"])
        keiyaku = "未契約" if ver == "追加" else "契約済"

    # 税区分：台帳は税抜で統一されているため、税込入力は税抜へ換算して保存する
    t1, t2 = st.columns([2, 1])
    zeikei = t1.radio("入力金額の税区分", ["税抜", "税込"], horizontal=True,
                      help="請求書・見積書の金額が税込の場合は『税込』を選ぶと、税抜に換算して保存します（台帳は税抜）。")
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
    edited["金額計"] = edited.apply(calc_amt, axis=1)       # 入力ベース（税抜/税込そのまま）
    edited["税抜金額"] = edited["金額計"].apply(to_net)      # 保存・集計は税抜に統一

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
    date_col = "請求日" if is_seikyu else "日付"
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
                        # 実績取込: 工種 / 区分 / 項目 / 金額 / 業者 / 請求日 / 支払区分
                        rows = [[r["工種"], r["区分"], r["項目"], int(round(r["税抜金額"])), gyosha, hiduke, shiharai]
                                for _, r in valid.iterrows()]
                        append_rows(sid, SHEET_JI, rows)
                    else:
                        # 見積取込: 工種 / 版 / 区分 / 項目 / 数量 / 単位 / 単価 / 金額 / 業者 / 日付 / 契約状況
                        rows = [[r["工種"], ver, r["区分"], r["項目"], to_num(r["数量"]) or "", r["単位"],
                                 int(round(to_net(r["単価"]))) or "", int(round(r["税抜金額"])), gyosha, hiduke, keiyaku]
                                for _, r in valid.iterrows()]
                        append_rows(sid, SHEET_MI, rows)
                    st.success(f"{koji} の「{target_sheet}」に {len(rows)} 行を税抜で追記しました。台帳（総合管理表など）が自動更新されます。")
                    clear_data_cache()
                    st.session_state.rows = default_rows()
                except Exception as e:
                    st.error(f"保存に失敗しました。共有と secrets 設定を確認してください。詳細: {e}")

# ────────────────────────────────────────────────────────
# 修正・取消タブ（見積取込 / 実績取込 を表で開いて行の編集・削除 → 全体を安全に上書き）
# ────────────────────────────────────────────────────────
def render_edit(koji, sid):
    st.subheader(f"🛠 {koji} のデータ修正・取消")
    st.caption("保存済みの取込データを表で開き、セルの修正や行の削除ができます。『上書き保存』でシート全体を置き換えます（誤読の差し替え・二重計上の取消に）。台帳シートは触りません。")

    label2sheet = {"見積取込": SHEET_MI, "実績取込": SHEET_JI}
    pick = st.selectbox("編集するデータ", list(label2sheet.keys()))
    sheet = label2sheet[pick]

    if st.button("🔄 最新に読み込み"):
        clear_data_cache()

    vals, hidx, err = read_sheet_raw(sid, sheet)
    if err:
        st.warning(f"「{sheet}」を読めませんでした: {err}")
        return
    if not vals or len(vals) <= hidx + 1:
        st.info("このシートには編集できるデータがありません。")
        return

    header = uniquify(vals[hidx])
    width = len(header)
    top = vals[:hidx + 1]  # タイトル行＋ヘッダ（保持する）
    data = [(row + [""] * width)[:width] for row in vals[hidx + 1:]]
    # 末尾の完全空行は表示から除く
    while data and all(str(c).strip() == "" for c in data[-1]):
        data.pop()
    df = pd.DataFrame(data, columns=header)

    st.markdown(f"#### {pick}（{len(df)} 行）— 行の削除・セルの修正ができます")
    edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=True)

    n_before, n_after = len(df), len(edited)
    if n_after != n_before:
        st.info(f"行数: {n_before} → {n_after}（{n_after - n_before:+d}）")

    st.warning("『上書き保存』はシート全体を現在の表の内容に置き換えます。保存の直前に自動で退避を取るので、直後なら『直前の状態に戻す』で1回だけ元に戻せます。")
    confirm = st.checkbox("内容を確認しました。上書き保存する")
    if st.button("上書き保存", type="primary", disabled=not confirm):
        try:
            body = [[("" if pd.isna(v) else v) for v in row] for row in edited.values.tolist()]
            snapshot_sheet(sid, sheet)  # 上書き前スナップショット（ワンステップ取り消し用）
            overwrite_ws(sid, sheet, top + body)
            clear_data_cache()
            st.success(f"「{sheet}」を上書き保存しました（{len(body)} 行）。直前の内容は「{BAK_PREFIX + sheet}」タブに退避しています。台帳が自動更新されます。")
        except Exception as e:
            st.error(f"保存に失敗しました。詳細: {e}")

    # ── ワンステップ取り消し（直前の退避から復元）──
    bak_vals, _, bak_err = read_sheet_raw(sid, BAK_PREFIX + sheet)
    if not bak_err and bak_vals:
        st.divider()
        st.caption(f"直前の上書き前の状態が「{BAK_PREFIX + sheet}」に保存されています。誤って上書きした場合はここから戻せます。")
        if st.button(f"⏪ 「{pick}」を直前の状態に戻す", key=f"restore_{sheet}"):
            try:
                overwrite_ws(sid, sheet, bak_vals)
                clear_data_cache()
                st.success(f"「{sheet}」を直前の状態に戻しました。台帳が自動更新されます。")
            except Exception as e:
                st.error(f"復元に失敗しました。詳細: {e}")

# ────────────────────────────────────────────────────────
st.title("見積・請求 取込")
st.caption("このアプリは入力（見積取込・実績取込への追記）専用です。予実・粗利・支払予定などの集計は、スプレッドシートの台帳（総合管理表・工種別内訳・支払予定 など）が自動で行います。")

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

tab_import, tab_edit = st.tabs(["📥 取込", "🛠 修正・取消"])
with tab_import:
    render_import(koji, sid, koshu)
with tab_edit:
    render_edit(koji, sid)
