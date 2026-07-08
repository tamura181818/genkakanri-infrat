"""
ファイル軽量化アプリ（Streamlit）— 画像で重くなった Excel / PDF を自動で軽くする

現場担当者がやること … ファイル(.xlsx か .pdf)を上げる → ボタンを押す → 軽くなったファイルを保存

【Excel(.xlsx)】
  中身のZIPにある画像(xl/media/)だけを縮小+再圧縮。表・数式・レイアウトはそのまま。
  写真系のPNG(透過なし)は自動でJPEGに変換して強力に圧縮。

【PDF】
  各ページの画像を指定解像度(DPI)まで縮小し再圧縮。文字レイヤーや押印は残します。
  「◯MB以下にする」を指定すると、その容量に収まるまで自動で調整します。
  ※ 圧縮は画質を落とすだけで、金額・文言・押印などの内容は一切書き換えません。

  pip install streamlit Pillow pymupdf
  streamlit run app.py
"""
import io
import os
import zipfile

import streamlit as st
from PIL import Image

# ══════════════════════════════════════════════════════════
#  Excel(.xlsx) 圧縮
# ══════════════════════════════════════════════════════════
TARGET_EXTS = {".png", ".jpg", ".jpeg"}


def _process_media(raw, ext, max_dim, quality, convert_png):
    """xlsx内の画像1枚を縮小・再圧縮。触らない場合は None。返り値:(bytes, 新拡張子)"""
    if ext not in TARGET_EXTS:
        return None
    try:
        im = Image.open(io.BytesIO(raw))
        im.load()
    except Exception:
        return None

    w, h = im.size
    if max(w, h) > max_dim:
        f = max_dim / float(max(w, h))
        im = im.resize((max(1, round(w * f)), max(1, round(h * f))), Image.LANCZOS)

    has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)

    if ext in (".jpg", ".jpeg"):
        buf = io.BytesIO()
        im.convert("RGB").save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
        return buf.getvalue(), ext
    if convert_png and not has_alpha:
        buf = io.BytesIO()
        im.convert("RGB").save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
        return buf.getvalue(), ".jpeg"
    buf = io.BytesIO()
    im.save(buf, "PNG", optimize=True)
    return buf.getvalue(), ".png"


def compress_xlsx(data, max_dim, quality, convert_png):
    """xlsx(bytes) を圧縮。返り値:(bytes, 変更した画像枚数)"""
    with zipfile.ZipFile(io.BytesIO(data)) as zin:
        names = zin.namelist()
        raws = {n: zin.read(n) for n in names}

    processed = {}
    rename_map = {}
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
            processed[name] = (name, raw)
            continue
        changed += 1
        if new_ext != ext:
            new_name = name[: -len(os.path.splitext(name)[1])] + new_ext
            rename_map[os.path.basename(name)] = os.path.basename(new_name)
            processed[name] = (new_name, new_bytes)
        else:
            processed[name] = (name, new_bytes)

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

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            fname, fbytes = processed[name]
            zout.writestr(fname, fbytes)

    result = out.getvalue()
    with zipfile.ZipFile(io.BytesIO(result)) as check:
        if check.testzip() is not None:
            raise RuntimeError("圧縮後のファイル検証に失敗しました。")
    return result, changed


# ══════════════════════════════════════════════════════════
#  PDF 圧縮
# ══════════════════════════════════════════════════════════
def compress_pdf(data, target_dpi, quality):
    """PDF(bytes) 内の画像を target_dpi 上限・quality で再圧縮。返り値:(bytes, 変更枚数)"""
    import fitz  # PyMuPDF

    doc = fitz.open(stream=data, filetype="pdf")
    changed = 0
    for page in doc:
        for img in page.get_images(full=True):
            xref = img[0]
            rects = page.get_image_rects(xref)
            if not rects:
                continue
            disp_w = max(r.width for r in rects)
            disp_h = max(r.height for r in rects)
            try:
                base = doc.extract_image(xref)
            except Exception:
                continue
            pw, ph = base["width"], base["height"]
            if not pw or not ph:
                continue
            scale = min((disp_w / 72 * target_dpi) / pw, (disp_h / 72 * target_dpi) / ph, 1.0)
            try:
                pil = Image.open(io.BytesIO(base["image"]))
                pil.load()
            except Exception:
                continue
            if scale < 0.98:
                pil = pil.resize((max(1, int(pw * scale)), max(1, int(ph * scale))), Image.LANCZOS)
            gray = pil.mode in ("L", "1") or base.get("colorspace") == 1
            pil = pil.convert("L") if gray else pil.convert("RGB")
            buf = io.BytesIO()
            pil.save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
            if len(buf.getvalue()) < base["size"]:
                try:
                    page.replace_image(xref, stream=buf.getvalue())
                    changed += 1
                except Exception:
                    pass

    # 画像を1枚も変更していない（＝文字だけ等）なら、再保存はしない。
    # PyMuPDFの再保存は最適化済みPDFを逆に膨らませることがあるため。
    if changed == 0:
        doc.close()
        return data, 0

    out = doc.tobytes(deflate=True, deflate_images=True, garbage=4, clean=True)
    doc.close()
    # 万一、元より大きくなったら元のまま返す（絶対に膨らませない）
    if len(out) >= len(data):
        return data, 0
    return out, changed


# 目標サイズに収めるための試行段階（画質優先→容量優先）
_PDF_STEPS = [(200, 82), (180, 78), (160, 72), (150, 68), (130, 65), (120, 60), (110, 55), (100, 50)]


def compress_pdf_to_target(data, target_bytes):
    """target_bytes 以下に収まるまで DPI/画質を段階的に下げる。
    返り値:(bytes, 変更枚数, 使った設定文字列 or None, 目標達成したか)"""
    # すでに目標以下なら何もしない
    if len(data) <= target_bytes:
        return data, 0, None, True

    best = data
    best_changed = 0
    best_setting = None
    for dpi, q in _PDF_STEPS:
        out, ch = compress_pdf(data, dpi, q)
        if len(out) < len(best):
            best, best_changed, best_setting = out, ch, f"{dpi}dpi / 画質{q}"
        # 元より小さくなっていて、かつ目標を満たしたら確定
        if best is not data and len(best) <= target_bytes:
            return best, best_changed, best_setting, True
    return best, best_changed, best_setting, len(best) <= target_bytes


# ══════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════
def human_size(n):
    x = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if x < 1024 or unit == "GB":
            return f"{x:.0f} {unit}" if unit == "B" else f"{x:,.1f} {unit}"
        x /= 1024


st.set_page_config(page_title="ファイル軽量化", layout="centered")
st.title("📉 ファイル軽量化（Excel / PDF）")
st.caption("画像で重くなった Excel・PDF を、見た目や内容はそのままに軽くします。")

with st.sidebar:
    st.header("PDF の設定")
    pdf_mode = st.radio(
        "やり方", ["目標サイズ以下にする（おすすめ）", "画質を自分で決める"], index=0,
    )
    if pdf_mode.startswith("目標"):
        target_mb = st.number_input(
            "目標サイズ（MB）以下にする", min_value=0.2, max_value=50.0, value=2.0, step=0.1,
            help="このサイズに収まるまで自動で画質を調整します。入札の上限に合わせてください。",
        )
        pdf_dpi = pdf_q = None
    else:
        target_mb = None
        pdf_dpi = st.select_slider(
            "解像度（DPI）", options=[100, 110, 120, 130, 150, 160, 180, 200], value=150,
            help="低いほど軽く。150前後でも書類の文字・押印は十分読めます。",
        )
        pdf_q = st.slider("画質（JPEG）", 40, 95, 70)

    st.divider()
    st.header("Excel の設定")
    xlsx_dim = st.select_slider(
        "画像の最大サイズ（長辺px）", options=[800, 1000, 1200, 1600, 2000, 2400, 3000], value=1600,
    )
    xlsx_q = st.slider("画質（Excel内画像）", 40, 95, 70, key="xq")
    xlsx_png = st.checkbox("写真系PNGをJPEGに変換して強力圧縮", value=True)

st.warning(
    "圧縮は画質を下げるだけで、金額・文言・押印などの**内容は書き換えません**。"
    "提出前に、押印や金額がはっきり読めるか必ずご確認ください。",
    icon="⚠️",
)

files = st.file_uploader(
    "ファイル（.pdf / .xlsx）を選ぶ（複数まとめてOK）",
    type=["pdf", "xlsx"], accept_multiple_files=True,
)
st.caption(
    "💡 このツールは**写真・スキャン画像で重くなったファイル**に効きます。"
    "文字だけの書類はもともと軽く、これ以上は小さくできません。"
)

if files and st.button("軽くする", type="primary"):
    results = []
    prog = st.progress(0.0)
    for i, f in enumerate(files):
        data = f.getvalue()
        ext = f.name.lower().rsplit(".", 1)[-1]
        rec = {"name": f.name, "before": len(data), "after": None, "data": None,
               "note": "", "error": None}
        try:
            with st.spinner(f"処理中… {f.name}"):
                if ext == "pdf":
                    if target_mb is not None:
                        out, ch, setting, ok = compress_pdf_to_target(
                            data, int(target_mb * 1024 * 1024))
                    else:
                        out, ch = compress_pdf(data, pdf_dpi, pdf_q)
                        setting = f"{pdf_dpi}dpi / 画質{pdf_q}"
                        ok = True
                    if len(out) >= len(data):
                        # 1バイトも減らなかった＝写真が無い/既に最適化済み
                        rec["note"] = (
                            "この書類は圧縮できる写真・画像が無い（または既に十分軽い）ため、"
                            "これ以上は小さくできませんでした。元のファイルをそのままお使いください。"
                        )
                        rec["ok"] = False
                    elif not ok:
                        rec["note"] = (
                            f"できるだけ圧縮しました（{setting} / 画像{ch}枚）が、"
                            "目標サイズには届きませんでした。"
                        )
                        rec["ok"] = False
                    else:
                        rec["note"] = f"設定: {setting} / 画像{ch}枚を圧縮"
                        rec["ok"] = True
                elif ext == "xlsx":
                    out, ch = compress_xlsx(data, xlsx_dim, xlsx_q, xlsx_png)
                    if len(out) >= len(data):
                        rec["note"] = (
                            "このExcelは圧縮できる画像が無い（または既に十分軽い）ため、"
                            "これ以上は小さくできませんでした。"
                        )
                        rec["ok"] = False
                    else:
                        rec["note"] = f"画像{ch}枚を圧縮"
                        rec["ok"] = True
                else:
                    raise RuntimeError("対応していない形式です。")
            rec["after"] = len(out)
            rec["data"] = out
        except Exception as e:
            rec["error"] = str(e)
        results.append(rec)
        prog.progress((i + 1) / len(files))
    st.session_state["results"] = results

for r in st.session_state.get("results", []):
    st.divider()
    if r["error"]:
        st.error(f"❌ {r['name']}：処理に失敗しました。詳細: {r['error']}")
        continue
    before, after = r["before"], r["after"]
    saved = before - after
    ratio = (saved / before * 100) if before else 0
    st.subheader(r["name"])
    c1, c2, c3 = st.columns(3)
    c1.metric("元のサイズ", human_size(before))
    c2.metric("圧縮後", human_size(after),
              delta=(f"-{human_size(saved)}" if saved > 0 else None), delta_color="inverse")
    c3.metric("削減率", f"{ratio:.0f}%")
    if r.get("note"):
        (st.caption if r.get("ok", True) else st.info)(r["note"])
    if saved <= 0:
        # 小さくできなかったので、ダウンロードは元ファイルと同じ。混乱を避けるため案内。
        st.caption("※ ダウンロードしても中身・容量は元と同じです。")
    base, ext = os.path.splitext(r["name"])
    mime = ("application/pdf" if ext.lower() == ".pdf"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button(
        "⬇ 軽くしたファイルをダウンロード",
        data=r["data"], file_name=f"{base}_軽量{ext}", mime=mime, key=f"dl_{r['name']}",
    )
