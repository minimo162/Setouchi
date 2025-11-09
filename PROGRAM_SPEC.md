# Excel Copilot 翻訳システム 現行仕様書

本書は `/workspaces/Setouchi` で維持されている Excel Copilot アプリケーションの最新構成・振る舞い・依存関係をまとめたものです。Flet 製デスクトップ UI、Playwright 経由で操作する M365 Copilot、xlwings を介した Excel 自動化という 3 レイヤーを束ね、翻訳作業と翻訳レビューをセル単位で支援します。

---

## 1. 提供モードと入出力

| モード | 主な目的 | 必須フィールド | 出力 | 実装ツール |
| --- | --- | --- | --- | --- |
| 翻訳（通常） | 参照資料を使わずに日本語→任意言語（デフォルト英語）へ翻訳 | `cell_range`, `translation_output_range` | 翻訳列（必要に応じて上書き） | `translate_range_without_references` |
| 翻訳（参照あり） | 原文・訳文両方の参照 URL を考慮した翻訳＋証跡 | `cell_range`, `translation_output_range`, (任意で `source_reference_urls`, `target_reference_urls`, `citation_output_range`) | 翻訳列＋説明／参照ペア列（幅は自動調整）。証跡は翻訳列または `citation_output_range` に配置 | `translate_range_with_references` |
| 翻訳チェック | 既存訳の評価と修正案・ハイライトの提示 | `source_range`, `translated_range`, `review_output_range`（3〜4 列） | 列 1=ステータス(OK/REVISE)、列 2=指摘メモ（日本語）、列 3=差分ハイライト文字列、列 4（任意）=修正案 | `check_translation_quality` |

フォーム UI は `MODE / SCOPE / OUTPUT / OPTIONS / REFERENCES` のタブで構成され、`CopilotMode` ごとに表示・バリデーション・サマリ生成が切り替わります。リスト型フィールドは改行またはカンマで区切って最大 3 件程度を想定し、空白は送信前に除去されます。

---

## 2. システム構成

### 2.1 デスクトップ UI（`desktop_app.py`）
- Flet を用い、上部ヒーロー（状態/統計/ショートカット）、左側コンテキストパネル（ブック・シート選択、同期ステータス、アクション）、タブ式フォーム、右側チャットタイムラインを備えた 3 カラムレイアウトを構築します。画面幅に応じてコンテキストをドロワー化し、ヒーロー内の数値カードやタイトルはレスポンシブに再配置されます。
- Excel のアクティブブック／シートは `_refresh_excel_context` とバックグラウンドポーリングスレッドで追跡し、候補ドロップダウンを常に最新化します。状態は `COPILOT_USER_DATA_DIR/setouchi_state.json` と `setouchi_logs` に永続化され、次回起動時に前回選択したブック・シート・フォーム値を復元します。
- フォーム送信時に `RequestMessage(USER_INPUT)` をエンキューし、UI 状態を `TASK_IN_PROGRESS` に遷移。進行中はプログレスリングとステータス文がアクションバーに表示され、停止ボタン（`RequestMessage(STOP)`）とブラウザリセットボタン（`RESET_BROWSER`）が活性化されます。
- チャットタイムラインは `ResponseType` ごとにカードスタイルを変え、LLM へ送ったプロンプトや Copilot 応答（`CHAT_PROMPT` / `CHAT_RESPONSE`）を含むすべてのログを時系列で残します。ログ保存／エクスポート、コマンドパレット、Excel フォーカス呼び出し (`_focus_excel_window`) といった補助操作もここから行えます。
- 起動時に `CopilotWorker` スレッドとレスポンス監視スレッドを生成し、UI スレッドからメッセージキュー越しに制御します。環境変数 `COPILOT_AUTOTEST_*` が設定されている場合は、指定のワークブック・シート・参照 URL・遅延・タイムアウトで自動テスト用フォーム送信を行い、終了後にアプリを自動終了します。

### 2.2 メッセージングとバックグラウンドワーカー（`excel_copilot/ui/messages.py`, `excel_copilot/ui/worker.py`）
- `RequestMessage`/`ResponseMessage` はシリアライズ可能な辞書構造で、UI ↔ ワーカー間の境界を明確化します。主なリクエスト種別は `USER_INPUT`, `STOP`, `QUIT`, `UPDATE_CONTEXT`, `RESET_BROWSER`。
- `CopilotWorker` は専用スレッドで常駐し、起動時に `BrowserCopilotManager` を初期化して `INITIALIZATION_COMPLETE` を通知します。モード変更 (`UPDATE_CONTEXT`) で利用可能ツール（翻訳／翻訳+参照／レビュー）を切り替え、ワーカー内 `stop_event` をすべてのツール関数に渡します。
- タスク実行時は `ExcelManager` をコンテキストマネージャとして開き `ExcelActions` を生成、`browser_manager` と `stop_event` を添えてツール関数を呼び出します。`ExcelConnectionError` や `ToolExecutionError` は UI へ `ERROR` レスポンスとして返却し、ユーザー停止は `UserStopRequested` を捕捉したうえでブラウザセッションのリセットを試みます。
- 進行ログは `ExcelActions.log_progress` → `ResponseType.STATUS` として UI にストリーミングされ、ツール終了時は `FINAL_ANSWER` → `END_OF_TASK` を送信して READY 状態へ戻します。停止・エラー後は `{"action": "focus_app_window"}` を付与した `INFO` を送出し、UI が Flet ウィンドウを前面に呼び戻します。

### 2.3 Copilot ブラウザ自動化（`excel_copilot/core/browser_copilot_manager.py`）
- Playwright の永続コンテキストを用いて Edge/Chrome プロファイル（`COPILOT_USER_DATA_DIR`）を再利用し、`COPILOT_BROWSER_CHANNELS` 優先順で起動します。`COPILOT_SUPPRESS_BROWSER_FOCUS` を有効にすると CDP でウィンドウを画面外へ移動し、ユーザーの操作を妨げないようにします。
- `ask(prompt, stop_event)` はチャット入力欄にプロンプトを貼り付け、送信ボタン／ショートカットを駆使して送信、最新メッセージの「コピー」ボタンを待って `pyperclip` から応答全文を取得します。Thought/Action フォーマットを返した場合は `Thought:` 以降のみを抽出。取得したプロンプト・応答は `set_chat_transcript_sink` で登録されたコールバック（UI 側チャットログ）に逐次中継されます。
- 途中停止は `request_stop()`（停止ボタンや Escape キー）で Copilot 側の生成を中断します。ユーザーが STOP を押した場合やセッションが汚染された場合は `reset_chat_session()` で GPT モード再選択 → 既存ページの再利用、失敗時は `restart()` でブラウザ自体を再起動します。
- `AppSettings`（`excel_copilot/config.py`）経由で各種タイムアウトやスローモード、レスポンス取得ディレイ (`COPILOT_RESPONSE_BUFFER_MS`) を環境変数で制御できます。

### 2.4 Excel 接続（`excel_copilot/core/excel_manager.py`）
- xlwings を介して既存の Excel プロセスに接続し、対象ブック／シートのアクティブ化・一覧取得・フォーカス制御を行います。接続リトライは既定 5 回（2 秒間隔）で、失敗時は `ExcelConnectionError` を送出します。
- 書き戻し対象ブック名は UI から `workbook_name` として渡され、タスク中は `ExcelManager` が `book`/`app` を保持。UI からのコンテキスト更新や自動テスト設定に応じて `activate_workbook`・`activate_sheet` が呼ばれます。

### 2.5 Excel 操作用ユーティリティ（`excel_copilot/tools/actions.py`）
- `ExcelActions` は読み書き／コピー／数式セットに加え、列幅・行高調整、折返し設定、リッチテキスト差分着色 (`apply_diff_highlight_colors`) を担います。macOS などリッチテキスト非対応環境では自動で `[ADD]/[DEL]` マーカーのみにフォールバックし、進捗ログで理由を説明します。
- すべての操作は `progress_callback` 経由で UI に逐次報告され、例外は `ToolExecutionError` としてラップされます。長文セルの自動折返しは `_preferred_column_width` と `_line_height` を基準に行単位で推定します。

### 2.6 ツール層（`excel_copilot/tools/excel_tools.py`）
- `translate_range_contents` が翻訳処理の中核で、両翻訳モードから再利用されます。セル範囲の正規化、既存値の再利用、UTF-16 コード単位によるバッチサイズ制御（`EXCEL_COPILOT_TRANSLATION_ITEMS_PER_REQUEST`, `EXCEL_COPILOT_TRANSLATION_UTF16_BUDGET`）、JSON 応答の頑健なパース、`stop_event` チェックを包括します。
- 参照付き翻訳では追加で `_extract_source_sentences_batch`（参照 URL から日本語引用を最大 10 件抽出）と `_pair_target_sentences_batch`（引用ごとにターゲット言語の原文を対応付け）の 2 ステップ Copilot 呼び出しを行い、その結果を `translation_context` に載せて本番翻訳をリクエストします。
- `check_translation_quality` はステータス／指摘／ハイライトの 3 列（+任意で修正列）を同期的に書き戻し、`highlight_output_range` が指定されていれば差分テキストとスタイル行列を生成します。AI 応答から `[ADD]/[DEL]` 付きハイライトが得られない場合は `_build_diff_highlight` で自前生成し、リッチテキストが使えないときはマーカー付テキストをセルに残します。
- そのほか `writetocell` や `copyrange` など単純作業用ラッパーも同ファイルに含まれており、将来的な ReAct エージェント向けに `excel_copilot/agent/react_agent.py` が残っています（現行 UI からは未使用）。

---

## 3. 処理フロー

### 3.1 共通シーケンス
1. **フォーム送信**: UI が入力値を検証し、モード・ツール名・引数を含む辞書をキューに投入。送信内容のサマリはチャットログへ `user` メッセージとして残ります。
2. **タスク受付**: `CopilotWorker` がペイロードを検証し、必要に応じて最新のブック／シート名で `ExcelManager` を初期化。参照するツール関数に `actions` / `browser_manager` / `sheet_name` / `stop_event` を差し込みます。
3. **ツール実行**: 各ツールが Excel からデータを読み出し、`BrowserCopilotManager.ask` を発行して結果を解析。進行状況は `STATUS`、Copilot とのやりとりは `CHAT_PROMPT` / `CHAT_RESPONSE` として即時反映されます。
4. **書き戻しと完了通知**: `ExcelActions.write_range` が必要な範囲だけをバッチ書き込みし、ツール結果文字列を `FINAL_ANSWER` として返却。ワーカーは `END_OF_TASK` で UI の状態遷移を指示し、Ready に戻します。

### 3.2 翻訳（通常）
- `rows_per_batch` は `EXCEL_COPILOT_TRANSLATION_ITEMS_PER_REQUEST` と UTF-16 予算を考慮して決定し、同一セル内容はキャッシュして重複呼び出しを回避します。既存訳があり日本語を含まない場合は再利用しつつ、必要に応じて上書きします。
- `translation_output_range` が不足している場合は自動で列幅を調整し、1 列／1 セルのブロックに再配置されます（調整内容はメッセージに記録）。
- 翻訳結果が空、原文コピー、日本語混在といった不正ケースはただちに `ToolExecutionError` で停止し、UI に原因を提示します。

### 3.3 翻訳（参照あり）
- 参照 URL はソース・ターゲット別に正規化／重複除去され、`source_reference_urls` と `target_reference_urls` のどちらか一方でも指定されていれば実行されます。行単位でリクエストを固定（`rows_per_batch=1`）し、引用抽出→参照ペア抽出→実翻訳の 3 段階を逐次実行します。
- `translation_output_range` はソース列あたり最低 12 列（訳文 1 + 説明 1 + 参照ペア 10 列）を確保できるよう自動調整され、各列ブロックに翻訳と将来的なコンテキスト列を割り当てます。
- `citation_output_range` は以下のレイアウトを認識します: 単一列（行単位でまとめて記述）、ソース列と同数（セル毎に 1 列）、列×2（説明列＋参照列）、列×3（訳文三つ組）。サポート外の列幅が指定された場合は翻訳出力列に証跡を内包する旨をログします。

### 3.4 翻訳チェック
- フォームで指定する `review_output_range` は 3 列（ステータス／指摘／ハイライト）または 4 列（+修正案）で、UI 側 ` _derive_review_output_ranges` が各列を `status_output_range`, `issue_output_range`, `highlight_output_range`, `corrected_output_range` に分割します。
- Copilot へのプロンプトにはステージ別指示（Action フェーズで必ず JSON を返すこと／Final フェーズで `Final Answer` をまとめること）が埋め込まれており、応答は ID, status, notes, before/after, highlighted_text などを含む配列として期待されています。`status` は `OK/PASS/GOOD` 系を受理、それ以外は自動的に `REVISE` と見なします。
- ハイライト列は元訳と修正版の diff を保持し、`ExcelActions.apply_diff_highlight_colors` にスタイル行列を渡すことでセル内の部分着色を試みます。配色は `EXCEL_COPILOT_ENABLE_RICH_DIFF_COLORS` で制御し、未対応プラットフォームでは `[ADD]...[/][DEL]...[/]` マーカーのみを残します。

---

## 4. 参照資料・証跡・ハイライト
- 参照抽出フローでは Copilot が返した引用文や参照ペアから重複を除去し、翻訳ログに「参照ペア[行-番号]: 原文 -> 参照訳」の形で出力します。翻訳列とは別に証跡を残したい場合は `citation_output_range` を設定するか、翻訳結果メッセージ内のメモを利用します。
- `check_translation_quality` が生成するハイライト列は、セル内テキストをそのまま差し替えるだけでなく、スタイル情報を `highlight_styles` に蓄積し、`supports_diff_highlight_colors()` が真の場合のみ Excel リッチテキスト API による部分着色を実施します。長大なセルやスタイル件数超過時は自動的にスキップし、理由をステータスログへ記録します。

---

## 5. 自動テストと運用補助
- 環境変数 `COPILOT_AUTOTEST_ENABLED` または `COPILOT_AUTOTEST_PROMPT` が指定されると、自動でフォーム値を構成し（既定参照 URL は `DEFAULT_AUTOTEST_*` で管理）、ワーカー初期化→Excel コンテキスト取得→フォーム送信→完了待ちまでを無人実行します。`COPILOT_AUTOTEST_DELAY` と `COPILOT_AUTOTEST_TIMEOUT` で開始ディレイとタイムアウトを設定でき、ログは `AUTOTEST:` プレフィックス付きで stdout に出力されます。
- UI 右上の「ログを書き出す」ボタンは `setouchi_logs` にタイムスタンプ付き JSONL を保存し、後続の不具合解析に活用できます。
- バックグラウンドの Excel ポーリングは 0.8 秒間隔でアクティブブックをサンプリングし、手動で Excel を切り替えた場合でも UI のワークブック／シート選択状態と同期を保ちます。

---

## 6. エラー処理・キャンセル・フォーカス制御
- すべてのツールは `stop_event` を随所で確認し、ユーザーが停止ボタンを押すと `UserStopRequested` を投げて処理を即座に抜けます。ワーカーはその後 `BrowserCopilotManager.request_stop()` → `reset_chat_session()`（必要なら `restart()`）を実行し、UI に停止完了メッセージを返します。
- Excel 接続や Copilot 側タイムアウトなどの致命的エラーは `ResponseType.ERROR` で詳細を通知し、UI 側は Ready 状態に戻さず `AppState.ERROR` で待機します。ユーザーが再度フォームを送れば復帰可能です。
- フォーカス関連では、UI から `focus_app_window` アクションを受けると Flet ウィンドウを前面に、`focus_excel_window` 操作で xlwings を通じて Excel 側を前面に表示します。これによりブラウザ自動化→Excel 操作間のコンテキスト切り替えを確実にします。

---

以上が現行の Excel Copilot 翻訳システム仕様です。新たなモードやフィールドを追加する際は、本書の各レイヤー（UI・ワーカー・Browser・ExcelActions・ツール）の責務分担とシーケンスを踏襲し、必要に応じてここへ追記してください。
