from pathlib import Path

path = Path('excel_copilot/tools/excel_tools.py')
old = """            extraction_prompt_sections: List[str] = [
                f"以下の日本語引用文 (Step 2) と原文 (source_text) に意味が対応する {target_language} 参照文を target_reference_urls から抽出し、対訳ペアを作成してください。",
                "",
                "手順:",
                "- 各 context_id について、source_sentences の各文に最も意味が近い {target_language} 文を最大6件まで抽出する。",
                "- 抽出する文は資料本文に実際に存在する文章をそのまま引用し、語尾・句読点・大文字小文字を保持する。",
                "- 推測や要約を返さず、該当文がない場合は {\"pairs\": []} を返す。",
                "- 同じ文を重複して返さない。",
                "- 目次・要約・脚注など本文以外のセクションは除外する。",
                "",
                "出力形式:",
                "- JSON配列のみを返し、各要素は {\"pairs\": [{\"source_sentence\": \"...\", \"target_sentence\": \"...\"}, ...]} 形式とする。",
                "- 対応する文が見つからない場合は {\"pairs\": []} を返す。",
                "",
                "items(JSON):",
                extraction_items_json,
            ]"""
new = """            extraction_prompt_sections: List[str] = [
                f"タスク: 以下の日本語引用文 (Step 2) と原文 (source_text) を読み、意味が対応する {target_language} の参照文を target_reference_urls から抽出して対訳ペアを作成してください。",
                "",
                "進め方:",
                "- context_id ごとに `source_sentences` の各文の主題・固有名詞・数値・文脈を手掛かりに、指定された target_reference_urls の本文で対応する {target_language} 文を探す。",
                "- 参照対象は提供済みの URL に限定し、見出しや段落を順番に確認して最も関連度の高い文を優先的に抽出する。外部検索や別ページへの移動は行わない。",
                "- 1つの `source_sentence` につき最大1文を選び、必要に応じて段落を文単位に分割して最も意味が合致する部分のみを引用する。",
                "- 引用する文は参照資料に記載された文字列をそのまま用い、意訳・要約・翻訳・語順変更を行わない。句読点や記号も原文どおりに保持する。",
                "- 引用文に含まれる脚注番号やURLなどの付加情報は除去し、本文のみを残す。",
                "- 対応する文が見つからない場合はペアを生成せず、context_id ごとに {\\"pairs\\": []} を返す。推測で文を作成しない。",
                "- `source_sentences` の順序を保ち、対応した {target_language} 文のみを返す。",
                "",
                "出力フォーマット:",
                "- JSON配列のみを返す。各要素は {\\"pairs\\": [{\\"source_sentence\\": \\\"...\\\", \\\"target_sentence\\\": \\\"...\\\"}, ...]} 形式とし、context_id の順序を守る。",
                "- 一致する {target_language} 文が見つからない場合は {\\"pairs\\": []} を返す。",
                "",
                "言語ポリシー:",
                "- ツール実行中の説明や思考 (AI thought など) も含め、すべて日本語で記述する。",
                "",
                "items(JSON):",
                extraction_items_json,
            ]"""
text = path.read_text(encoding='utf-8')
if old not in text:
    raise SystemExit('old block not found')
path.write_text(text.replace(old, new), encoding='utf-8')
