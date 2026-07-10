"""
Gemini 呼び出しラッパ。

- API キーは環境変数 GEMINI_API_KEY から取得（既存 app.py と同方針）。
- structured output（response_schema）で JSON を強制し、パース失敗を減らす。
- google-generativeai 未インストール / キー未設定でも import は通り、
  実際に呼んだ時点で分かりやすいエラーを出す（scripts 単体テストのため）。

モデルは設計書 §4 に従い Flash / Flash-Lite を切替可能。
新規 API キーでは 'gemini-2.5-flash' 直指定が 404 になる場合があるため、
既定は安定エイリアスの 'gemini-flash-latest'（Flash-Lite は
'gemini-flash-lite-latest'）とする。GEMINI_MODEL で上書き可。
"""
from __future__ import annotations

import json
import os

DEFAULT_MODEL = os.environ.get("GEMINI_MODEL", "gemini-flash-latest")


class GeminiUnavailable(RuntimeError):
    pass


def _configure():
    key = os.environ.get("GEMINI_API_KEY")
    if not key:
        raise GeminiUnavailable(
            "GEMINI_API_KEY が未設定です。`export GEMINI_API_KEY=...` を設定してください。"
        )
    try:
        import google.generativeai as genai
    except ImportError as e:
        raise GeminiUnavailable(
            "google-generativeai が未インストールです。`pip install google-generativeai`"
        ) from e
    genai.configure(api_key=key)
    return genai


def call_vision_json(image_path: str, prompt: str, schema: dict,
                     model_name: str = DEFAULT_MODEL) -> dict:
    """画像 1 枚 + プロンプトを渡し、schema 準拠の JSON(dict) を返す。"""
    genai = _configure()
    with open(image_path, "rb") as f:
        image_bytes = f.read()

    model = genai.GenerativeModel(model_name)
    resp = model.generate_content(
        [{"mime_type": "image/jpeg", "data": image_bytes}, prompt],
        generation_config={
            "response_mime_type": "application/json",
            "response_schema": schema,
            "temperature": 0.0,
        },
    )
    return _parse(resp)


def _parse(resp) -> dict:
    txt = (getattr(resp, "text", "") or "").strip()
    txt = txt.replace("```json", "").replace("```", "").strip()
    if not txt:
        return {}
    return json.loads(txt)
