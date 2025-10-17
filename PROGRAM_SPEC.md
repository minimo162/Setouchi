# Excel Copilot 翻訳システム仕様書

本ドキュメントは、Excel 上で稼働する翻訳支援アプリケーションの最新構成と挙動をまとめたものです。現在はフォーム経由でリクエストを受け付ける UI を標準とし、従来の自由入力チャットは閲覧専用に縮退しています。

## 目的と機能

- Excel の指定セル範囲を日本語から英語（他言語へも拡張可能）に翻訳する。
- 参照資料（URL やファイル）を利用して訳語統一を支援する。
- 既存翻訳の品質レビューを行い、指摘内容と修正案を提示する。
- 翻訳結果やレビュー結果を Excel に書き戻し、進捗を UI 上で可視化する。

## アーキテクチャ概要

| レイヤー | 役割 | 主なモジュール |
| --- | --- | --- |
| デスクトップ UI (Flet) | フォーム入力・チャット閲覧 UI を提供し、ユーザー入力をキューへ送信する。結果やステータスをタイムライン表示する。 | `desktop_app.py`, `excel_copilot/ui/chat.py`, `excel_copilot/ui/theme.py` |
| ワーカースレッド | Playwright 経由で Copilot セッションを維持し、フォームで受け取った構造化リクエストをツールへ橋渡しする。 | `excel_copilot/ui/worker.py`, `excel_copilot/ui/messages.py` |
| ツール層 | 翻訳・レビューなどの Excel 操作を実装し、Copilot とのプロンプト往復を担う。 | `excel_copilot/tools/excel_tools.py`, `excel_copilot/tools/actions.py` |
| コアサービス | Excel 接続（xlwings）と Copilot ブラウザ自動化（Playwright）、共通例外と設定。 | `excel_copilot/core/excel_manager.py`, `excel_copilot/core/browser_copilot_manager.py`, `excel_copilot/core/exceptions.py`, `excel_copilot/config.py` |

## 実行フロー（フォーム版）

1. **アプリ起動 (`desktop_app.py`)**  
   Flet アプリを初期化し、`CopilotWorker` をバックグラウンドで起動する。チャット履歴の表示とフォーム UI を構築する。
2. **ワーカー初期化 (`CopilotWorker._initialize`)**  
   `BrowserCopilotManager` を起動して Copilot Web チャットへ接続し、モードに応じたツール関数を読み込む。
3. **フォーム送信 (`CopilotApp._submit_form`)**  
   UI で入力された値を検証し、モード・ツール名・引数から成る JSON ペイロードを `RequestQueue` へ投入する。送信内容はチャットログにサマリとして記録される。
4. **構造化タスクの実行 (`CopilotWorker._run_structured_task`)**  
   受け取ったペイロードを検証し、該当ツールに必要な引数をバインドして実行する。進捗はステータスメッセージとして UI へ返す。
5. **Excel への書き戻し**  
   ツール階層が Copilot からの応答をもとに Excel を更新し、その結果メッセージをワーカー経由で UI に返却する。
6. **チャット表示・状態更新**  
   UI は Final Answer をチャットログへ反映し、タスク状態を READY へ戻す。停止ボタンやモードカードの状態も合わせて更新される。

## 今後の改善候補

- フォーム定義を外部 JSON として切り出し、モード追加時の拡張性を高める。
- 代表的な入力パターンに対する自動テスト（フォーム検証、Excel への反映テスト）を追加する。
- 参照ファイル添付（ローカルパス指定）の UX 向上やドラッグ&ドロップ対応を検討する。
