#!/usr/bin/env python3
"""受信FAX (PDF) をClaude APIで分類し、カテゴリ別フォルダに振り分ける。

使い方:
    python sort_fax.py             # ./inbox のPDFを処理
    python sort_fax.py path/to/dir # 指定ディレクトリのPDFを処理

環境変数:
    ANTHROPIC_API_KEY  Claude APIキー (必須)
"""
from __future__ import annotations

import base64
import json
import shutil
import sys
from pathlib import Path

import anthropic

MODEL = "claude-opus-5"

CATEGORIES = [
    "注文書",
    "請求書",
    "見積依頼",
    "納品書",
    "問い合わせ",
    "その他",
]

SYSTEM_PROMPT = (
    "あなたは日本企業で受信したFAXを分類するアシスタントです。\n"
    "添付PDFの内容を読み取り、最も適切なカテゴリを1つ選んでください。\n\n"
    "カテゴリ:\n" + "\n".join(f"- {c}" for c in CATEGORIES) + "\n\n"
    "判断に迷う場合は「その他」を選んでください。"
)

OUTPUT_SCHEMA = {
    "type": "object",
    "properties": {
        "category": {"type": "string", "enum": CATEGORIES},
        "reason": {"type": "string", "description": "分類根拠を1文で"},
    },
    "required": ["category", "reason"],
    "additionalProperties": False,
}


def classify_pdf(client: anthropic.Anthropic, pdf_path: Path) -> dict:
    pdf_b64 = base64.standard_b64encode(pdf_path.read_bytes()).decode("utf-8")

    response = client.messages.create(
        model=MODEL,
        max_tokens=1024,
        system=[
            {
                "type": "text",
                "text": SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {"type": "text", "text": "このFAXを分類してください。"},
                ],
            }
        ],
        output_config={"format": {"type": "json_schema", "schema": OUTPUT_SCHEMA}},
    )

    text = next(b.text for b in response.content if b.type == "text")
    return json.loads(text)


def main(inbox: Path, sorted_root: Path) -> int:
    if not inbox.is_dir():
        print(f"入力フォルダがありません: {inbox}", file=sys.stderr)
        return 1

    pdfs = sorted(inbox.glob("*.pdf"))
    if not pdfs:
        print(f"PDFが見つかりません: {inbox}")
        return 0

    client = anthropic.Anthropic()
    errors = 0

    for pdf in pdfs:
        try:
            result = classify_pdf(client, pdf)
        except Exception as e:
            print(f"[ERROR] {pdf.name}: {e}", file=sys.stderr)
            errors += 1
            continue

        category = result["category"]
        dest_dir = sorted_root / category
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / pdf.name
        shutil.move(str(pdf), str(dest))
        print(f"[{category}] {pdf.name} — {result['reason']}")

    print(f"\n完了: {len(pdfs) - errors}/{len(pdfs)} 件")
    return 1 if errors else 0


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("inbox")
    sys.exit(main(target, Path("sorted")))
