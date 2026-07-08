"""
Excel 画像圧縮アプリ（Streamlit）— 画像入りの重い .xlsx を自動で軽くする

現場担当者がやること … Excel(.xlsx)を上げる → 「圧縮する」を押す → 軽くなったファイルを保存

しくみ:
  .xlsx は中身が ZIP で、貼り付けた画像は xl/media/ に入っています。
  その画像だけを「縮小 + 再圧縮」して詰め直します。表や数式・レイアウトはそのまま。
  ・大きすぎる画像は指定サイズまで縮小
  ・写真系のPNG（透過なし）は自動でJPEGに変換（ここで一番容量が減ります）
  ・透過が必要なPNGはPNGのまま最適化（見た目が崩れないよう配慮）

  pip install streamlit Pillow
  streamlit run app.py
"""
import io
import os
import zipfile

import streamlit as st
from PIL import Image

# xl/media/ 内で処理対象にする画像形式（emf/wmf などのベクタや gif は安全のため触らない）
TARGET_EXTS = {".png", ".jpg", ".jpeg"}


def _process_media(raw, ext, max_dim, quality, convert_png):
    """1枚の画像を縮小・再圧縮する。触らない場合は None を返す。
    返り値: (新しいバイト列, 新しい拡張子)
    """
    if ext not in TARGET_EXTS:
        return None
    try:
        im = Image.open(io.BytesIO(raw))
        im.load()
    except Exception:
        return None  # 読めない画像はそのまま

    # ── 縮小（縦横の長い方を max_dim に収める）──
    w, h = im.size
    if max(w, h) > max_dim:
        f = max_dim / float(max(w, h))
        im = im.resize((max(1, round(w * f)), max(1, round(h * f))), Image.LANCZOS)

    has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)

    # ── 既存JPEG: 画質を落として再圧縮 ──
    if ext in (".jpg", ".jpeg"):
        buf = io.BytesIO()
        im.convert("RGB").save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
        return buf.getvalue(), ext

    # ── PNG（透過なし）を JPEG に変換（容量が最も減る）──
    if convert_png and not has_alpha:
        buf = io.BytesIO()
        im.convert("RGB").save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
        return buf.getvalue(), ".jpeg"

    # ── 透過PNG など: PNGのまま最適化（縮小の効果は反映される）──
    buf = io.BytesIO()
    im.save(buf, "PNG", optimize=True)
    return buf.getvalue(), ".png"


def compress_xlsx(data, max_dim, quality, convert_png):
    """xlsx(bytes) を受け取り、画像を圧縮した xlsx(bytes) を返す。
    返り値: (新しいバイト列, 変更した画像枚数)
    """
    with zipfile.ZipFile(io.BytesIO(data)) as zin:
        names = zin.namelist()
        raws = {n: zin.read(n) for n in names}

    processed = {}       # 元のファイル名 -> (最終ファイル名, バイト列)
    rename_map = {}      # 拡張子が変わった画像の basename 対応（png -> jpeg）
    changed = 0

    for name in names:
        raw = raws[name]
        low = name.lower()
        if not low.startswith("xl/media/"):
            processed[name] = (name, raw)
            continue

        ext = os.path.splitext(name)[1].lower()
        res = _process_media(raw, ext, max_dim, quality, convert_png)
        if res is None:
            processed[name] = (name, raw)
            continue

        new_bytes, new_ext = res
        if len(new_bytes) >= len(raw):
            processed[name] = (name, raw)  # 小さくならないなら元のまま
            continue

        changed += 1
        if new_ext != ext:
            new_name = name[: -len(os.path.splitext(name)[1])] + new_ext
            rename_map[os.path.basename(name)] = os.path.basename(new_name)
            processed[name] = (new_name, new_bytes)
        else:
            processed[name] = (name, new_bytes)

    # ── 拡張子が変わった画像がある場合、参照(rels)と Content_Types を書き換える ──
    if rename_map:
        for name in list(processed.keys()):
            low = name.lower()
            if not (low.endswith(".rels") or low == "[content_types].xml"):
                continue
            fname, fbytes = processed[name]
            try:
                text = fbytes.decode("utf-8")
            except UnicodeDecodeError:
                continue
            for old, new in rename_map.items():
                text = text.replace(old, new)
            if low == "[content_types].xml" and 'Extension="jpeg"' not in text:
                text = text.replace(
                    "</Types>", '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>'
                )
            processed[name] = (fname, text.encode("utf-8"))

    # ── 元の順番のまま書き出す ──
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            fname, fbytes = processed[name]
            zout.writestr(fname, fbytes)

    result = out.getvalue()
    # 壊れていないか最終チェック
    with zipfile.ZipFile(io.BytesIO(result)) as check:
        if check.testzip() is not None:
            raise RuntimeError("圧縮後のファイル検証に失敗しました。")
    return result, changed


def human_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:,.1f} {unit}"
        n /= 1024


# ────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel 画像圧縮", layout="centered")
st.title("📉 Excel 画像圧縮")
st.caption("画像を貼って重くなった Excel(.xlsx) を、レイアウトはそのままに軽くします。")

with st.sidebar:
    st.header("圧縮の設定")
    max_dim = st.select_slider(
        "画像の最大サイズ（長辺・ピクセル）",
        options=[800, 1000, 1200, 1600, 2000, 2400, 3000],
        value=1600,
        help="写真をこのサイズまで縮小します。小さいほど軽くなります。印刷用途なら 2000 前後がおすすめ。",
    )
    quality = st.slider(
        "画質（JPEG）", min_value=40, max_value=95, value=70,
        help="低いほど軽く、荒くなります。70前後が実用的なバランスです。",
    )
    convert_png = st.checkbox(
        "写真系のPNGをJPEGに変換して強力圧縮", value=True,
        help="容量が最も減ります。透過（背景ぬき）が必要な画像は自動でPNGのまま残します。",
    )

files = st.file_uploader(
    "Excelファイル（.xlsx）を選ぶ（複数まとめてOK）",
    type=["xlsx"], accept_multiple_files=True,
)
st.info("※ 古い形式の .xls には対応していません。Excelで「.xlsx」として保存し直してからお使いください。")

if files and st.button("圧縮する", type="primary"):
    results = []
    prog = st.progress(0.0)
    for i, f in enumerate(files):
        data = f.getvalue()
        try:
            with st.spinner(f"圧縮中… {f.name}"):
                out, changed = compress_xlsx(data, max_dim, quality, convert_png)
            results.append({
                "name": f.name, "before": len(data), "after": len(out),
                "changed": changed, "data": out, "error": None,
            })
        except Exception as e:
            results.append({
                "name": f.name, "before": len(data), "after": None,
                "changed": 0, "data": None, "error": str(e),
            })
        prog.progress((i + 1) / len(files))
    st.session_state["results"] = results

for r in st.session_state.get("results", []):
    st.divider()
    if r["error"]:
        st.error(f"❌ {r['name']}：圧縮に失敗しました。詳細: {r['error']}")
        continue

    before, after = r["before"], r["after"]
    saved = before - after
    ratio = (saved / before * 100) if before else 0
    st.subheader(r["name"])
    c1, c2, c3 = st.columns(3)
    c1.metric("元のサイズ", human_size(before))
    c2.metric("圧縮後", human_size(after), delta=f"-{human_size(saved)}", delta_color="inverse")
    c3.metric("削減率", f"{ratio:.0f}%")

    if r["changed"] == 0:
        st.info("画像が見つからないか、これ以上は軽くできませんでした（すでに軽い可能性があります）。")
    else:
        st.caption(f"{r['changed']} 枚の画像を圧縮しました。")

    base, _ = os.path.splitext(r["name"])
    st.download_button(
        "⬇ 軽くしたファイルをダウンロード",
        data=r["data"],
        file_name=f"{base}_軽量.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_{r['name']}",
    )
