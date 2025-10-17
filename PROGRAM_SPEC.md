# Excel Copilot 翻訳システム仕様書

本ドキュメントは、Excel 上で稼働する翻訳支援アプリケーションの最新構成と挙動をまとめたものです。現在はフォーム経由でリクエストを受け付ける UI を標準とし、従来の自由入力チャットは閲覧専用に縮退しています。

## 目的と機能

- Excel の指定セル範囲を日本語から英語（他言語にも拡張可能）へ翻訳する。
- 参照資料（URL やファイル）を活用し、訳語の統一と証跡を補強する。
- 既存翻訳の品質レビューを行い、指摘内容・修正案・ハイライトを提示する。
- 翻訳結果やレビュー結果を Excel に書き戻し、進捗を UI 上で可視化する。

## アーキテクチャ概要

| レイヤー | 役割 | 主なモジュール |
| --- | --- | --- |
| デスクトップ UI (Flet) | フォーム入力とチャット閲覧 UI を提供し、ユーザー入力をキューへ送信する。結果やステータスをタイムライン表示する。 | `desktop_app.py`, `excel_copilot/ui/chat.py`, `excel_copilot/ui/theme.py` |
| ワーカースレッド | Playwright 経由で Copilot セッションを維持し、フォームで受け取った構造化リクエストをツールへ橋渡しする。 | `excel_copilot/ui/worker.py`, `excel_copilot/ui/messages.py` |
| ツール層 | 翻訳・レビューなどの Excel 操作を実装し、Copilot とのプロンプト往復を担う。 | `excel_copilot/tools/excel_tools.py`, `excel_copilot/tools/actions.py` |
| コアサービス | Excel 接続（xlwings）と Copilot ブラウザ自動化（Playwright）、共通例外と設定。 | `excel_copilot/core/excel_manager.py`, `excel_copilot/core/browser_copilot_manager.py`, `excel_copilot/core/exceptions.py`, `excel_copilot/config.py` |

## 実行フロー（フォーム版）

1. **アプリ起動（`desktop_app.py`）**  
   Flet アプリを初期化し、`CopilotWorker` をバックグラウンド起動する。チャット履歴とフォーム UI を構築し、初期状態を READY に設定する。
2. **ワーカー初期化（`CopilotWorker._initialize`）**  
   `BrowserCopilotManager` を起動して Copilot Web チャットへ接続し、モードに応じたツール関数を読み込む。初期化完了までステータスメッセージを UI に通知する。
3. **フォーム送信（`CopilotApp._submit_form`）**  
   必須入力・数値範囲・参照 URL の有無を検証し、モード・ツール名・引数から成る JSON ペイロードを `RequestQueue` へ投入する。翻訳モードでは原文／訳文向けの参照 URL のみを受け付け、共通参照・引用出力・バッチ行数といった旧フィールドは扱わない。参照入力はフォーム内の専用セクションにまとめ、不要時は空欄のまま送信できる。ユーザーには要約をチャットログとして表示する。
4. **構造化タスク実行（`CopilotWorker._run_structured_task`）**  
   ペイロードを検証し、対象ブック／シート・引数を整えて該当ツールを実行する。進捗メッセージや最終結果を UI に返却する。
5. **Excel 書き戻し**  
   ツール層が Copilot 応答をもとに Excel を更新し、その結果メッセージをワーカー経由で UI に返す。レビュー系タスクではフォームで指定された `highlight_output_range` に差分ハイライトを描画し、リッチテキスト非対応環境ではマーカー付きテキストにフォールバックする。停止要求があった場合は途中で処理を中断し、ブラウザセッションをリセットする。
6. **UI 更新と終了処理**  
   Final Answer をチャットログに追加し、タスク状態を READY へ戻す。停止ボタンやモードカードの状態も現在のステータスに合わせて更新する。

## 差分ハイライト

- レビュー用ツール（`check_translation_quality`）は AI 応答に含まれる `[ADD]...` と `[DEL]...` マーカーからセル単位の差分スパンを生成し、`highlight_output_range` に最新訳文を下敷きとして書き戻す。
- `ExcelActions.apply_diff_highlight_colors` が加筆箇所を青（`#1565C0`）、削除箇所を赤（`#C62828`）でセル内文字に着色し、プラットフォームがリッチテキスト非対応の場合はマーカーのみを残す。
- `highlight_text_differences` ユーティリティを使うことで、任意の 2 範囲間でも同等のハイライトを生成でき、レビュー以外のシナリオでも視覚的フィードバックを提供できる。

## 今後の改善候補

- フォーム定義を外部 JSON 等へ切り出し、モード追加やフィールド変更をコード修正なしで行えるようにする。
- フォーム入力と Excel 書き戻しを対象とした自動テストを整備し、回帰検証の確実性を高める。
- 参照資料の入力 UX を改善し、ドラッグ＆ドロップや最近使用したリンクの候補表示を追加検討する。
