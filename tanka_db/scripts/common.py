"""
共通ユーティリティ — パス解決 / 設定読込 / 文字正規化 / 単位・労務判定。

設計書 §1〜§4 に対応。パイプライン各スクリプト(01〜05)から import して使う。
外部依存は PyYAML のみ（pandas 等は各スクリプトが必要に応じて import）。
"""
from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

import yaml

# ── パス ─────────────────────────────────────────────
# scripts/ の 1 つ上が project ルート（tanka_db/）。
ROOT = Path(__file__).resolve().parent.parent
INPUT_PDFS = ROOT / "input_pdfs"
CONFIG = ROOT / "config"
WORK = ROOT / "work"
OUTPUT = ROOT / "output"

LABOR_EXCLUDE_YAML = CONFIG / "labor_exclude.yaml"
UNIT_NORMALIZE_YAML = CONFIG / "unit_normalize.yaml"


# ── work_id ──────────────────────────────────────────
def work_id_from_pdf(pdf_path: str | Path) -> str:
    """PDF ファイル名から work_id を生成（拡張子を除いた stem）。

    例: '01_R7国道357号塩浜立体海側橋梁上部その3工事.pdf'
        -> '01_R7国道357号塩浜立体海側橋梁上部その3工事'
    """
    return Path(pdf_path).stem


def work_dir(work_id: str) -> Path:
    return WORK / work_id


def pages_dir(work_id: str) -> Path:
    return work_dir(work_id) / "pages"


# ── 設定読込 ─────────────────────────────────────────
def load_yaml(path: str | Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def load_labor_exclude(path: str | Path = LABOR_EXCLUDE_YAML) -> dict:
    """labor_exclude.yaml を読み、判定しやすい形へ整形して返す。

    戻り値:
      {
        "occupations": {正規化済み職種名, ...},   # 全カテゴリを平坦化した集合
        "labor_units": ["人", "人日"],
        "synonyms": [("橋りょう","橋梁"), ...],
      }
    """
    raw = load_yaml(path)
    policy = raw.get("match_policy", {}) or {}
    norm = policy.get("normalize", {}) or {}

    synonyms = [tuple(pair) for pair in (norm.get("synonyms") or [])]

    occ_set: set[str] = set()
    for _category, names in (raw.get("occupations") or {}).items():
        for name in names or []:
            occ_set.add(normalize_name(name, synonyms))

    return {
        "occupations": occ_set,
        "labor_units": list(policy.get("labor_units") or ["人", "人日"]),
        "synonyms": synonyms,
    }


def load_unit_normalizer(path: str | Path = UNIT_NORMALIZE_YAML) -> dict:
    """unit_normalize.yaml を読み、別名->正規形 の逆引き辞書を返す。"""
    raw = load_yaml(path)
    alias_to_canon: dict[str, str] = {}
    for canon, aliases in (raw.get("aliases") or {}).items():
        alias_to_canon[_pre_normalize_unit(canon)] = canon
        for a in aliases or []:
            alias_to_canon[_pre_normalize_unit(a)] = canon
    return alias_to_canon


# ── 文字正規化 ───────────────────────────────────────
def to_halfwidth(s: str) -> str:
    """全角英数記号・全角スペースを半角へ（NFKC）。"""
    return unicodedata.normalize("NFKC", s)


def kana_to_fullwidth(s: str) -> str:
    """半角カナを全角カナへ（NFKC は半角カナも全角化するため実質これで足りる）。"""
    return unicodedata.normalize("NFKC", s)


def unify_parentheses(s: str) -> str:
    return s.replace("（", "(").replace("）", ")")


def strip_all_spaces(s: str) -> str:
    return re.sub(r"[\s　]+", "", s)


def apply_synonyms(s: str, synonyms: list[tuple[str, str]]) -> str:
    """synonyms の各ペアを 2 番目の表記に寄せる（例: 橋りょう->橋梁）。"""
    for pair in synonyms:
        if len(pair) < 2:
            continue
        variant, canon = pair[0], pair[1]
        s = s.replace(variant, canon)
    return s


def normalize_name(name: str, synonyms: list[tuple[str, str]] | None = None) -> str:
    """労務判定・辞書照合用の名称正規化（辞書 match_policy に準拠）。"""
    if name is None:
        return ""
    s = str(name)
    s = to_halfwidth(s)
    s = unify_parentheses(s)
    s = strip_all_spaces(s)
    if synonyms:
        s = apply_synonyms(s, synonyms)
    return s


def to_num(x):
    """'1,234.5' / '¥3,815' / 全角数字 などを float へ。空・不能は None。"""
    if x is None:
        return None
    s = to_halfwidth(str(x))
    s = re.sub(r"[^\d.\-]", "", s)
    if s in ("", "-", ".", "--"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


# ── 単位正規化 ───────────────────────────────────────
def _pre_normalize_unit(u: str) -> str:
    s = to_halfwidth(str(u or ""))
    s = strip_all_spaces(s)
    s = unify_parentheses(s)
    return s


def normalize_unit(unit, alias_to_canon: dict[str, str]) -> tuple[str, bool]:
    """単位を正規形へ。戻り値 (正規形, known)。未知なら前処理後の値と False。"""
    key = _pre_normalize_unit(unit)
    if key == "":
        return "", True
    if key in alias_to_canon:
        return alias_to_canon[key], True
    return key, False


# ── 労務判定（三段構え。設計書 §1 の判定ロジック） ──
def judge_labor(row: dict, labor_cfg: dict) -> tuple[str, str]:
    """1 行を 保持/除外/要レビュー に分類する（LLM 判定前の決定的ロジック）。

    row は 03_extract の出力 1 行（キー: 名称, 規格, 単位, 数量, 単価, 金額,
    摘要, is_reference, labor_type）。

    戻り値 (判定, 理由):
      判定 = "保持" | "除外" | "要レビュー"
    LLM 由来の labor_type を第 3 段の材料として利用する。
    """
    name = row.get("名称", "") or ""
    unit_raw = row.get("単位", "") or ""
    tanka = to_num(row.get("単価"))
    tekiyo = str(row.get("摘要", "") or "")
    is_ref = bool(row.get("is_reference", False))

    norm_name = normalize_name(name, labor_cfg.get("synonyms"))
    norm_unit = _pre_normalize_unit(unit_raw)
    labor_units = labor_cfg.get("labor_units", ["人", "人日"])
    occ = labor_cfg.get("occupations", set())

    kingaku = to_num(row.get("金額"))

    # 段1: 号参照フィルタ / 集計行フィルタ → 除外（中間集計であり最下層でない）
    if is_ref or re.search(r"(単-|内-)", tekiyo):
        return "除外", "号参照(中間集計)"
    if tanka is None and norm_unit == "式":
        return "除外", "式・単価空欄(費目集計行)"
    # 単価も金額も無い行は最下層単価ではない（見出し・区切り・空行）→ 除外
    if tanka is None and kingaku is None:
        return "除外", "非単価行(見出し/空欄)"

    # 段2: 除外辞書 × 単位併用判定（設計書 §1-2）
    #   純粋労務は 人/人日 単位が原則。過剰除去を防ぐため、施工単位(m2等)を
    #   伴う一致は「保持側に倒して要レビュー」とする。
    unit_is_labor = norm_unit in labor_units
    unit_is_empty_or_shiki = norm_unit in ("", "式")
    is_exact = norm_name in occ
    is_partial = any(term and term in norm_name for term in occ)

    #   (a) 名称が辞書に完全一致
    if is_exact:
        if unit_is_labor or unit_is_empty_or_shiki:
            return "除外", "純粋労務(辞書完全一致)"
        # 完全一致でも施工単位 → 保持側+要レビュー（例: 塗装工 /m2 は稀だが要確認）
        return "要レビュー", "辞書完全一致だが施工単位のため保持側"
    #   (b) 名称が辞書語を含み、かつ単位が労務単位 → 除外
    if is_partial and unit_is_labor:
        return "除外", "純粋労務(辞書部分一致+労務単位)"
    #   名称が辞書語を含むが単位が施工単位 → 保持側へ倒して要レビュー
    if is_partial:
        return "要レビュー", "辞書部分一致だが施工単位のため保持側"

    # 段3: LLM 分類の取り込み
    labor_type = row.get("labor_type")
    if labor_type == "pure_labor":
        # LLM が純粋労務と判断。単位も労務単位なら確定除外、そうでなければレビュー。
        if norm_unit in labor_units:
            return "除外", "純粋労務(LLM+労務単位)"
        return "要レビュー", "LLMが純粋労務だが単位が非労務"
    if labor_type == "unknown":
        return "要レビュー", "LLM判別不能"

    # ここまで来たら labor_type == "work_unit" もしくは未指定 → 保持
    return "保持", "労務を内包した施工単価"


# ── jsonl I/O ────────────────────────────────────────
def read_jsonl(path: str | Path) -> list[dict]:
    rows = []
    p = Path(path)
    if not p.exists():
        return rows
    with open(p, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows


def write_jsonl(path: str | Path, rows: list[dict]) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
