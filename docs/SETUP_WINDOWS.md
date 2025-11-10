# Windows セットアップ手順

1. **Rust / Cargo**
   - https://rustup.rs/ から `rustup-init.exe` を取得し、64bit (x86_64) デフォルト設定でインストール。
   - インストール後、`rustup update` を実行し最新化。

2. **Tauri CLI**
   - `cargo install tauri-cli --version ^2.0.0` を実行。

3. **Python & 依存ライブラリ**
   - `python -m pip install --upgrade pip`
   - `pip install xlwings playwright pyperclip`
   - `playwright install msedge`

4. **プロジェクトのビルド/実行**
   - リポジトリ直下で `powershell -ExecutionPolicy Bypass -File scripts/dev.ps1` を実行すると Tauri Dev サーバーが起動します。
   - リリースビルドは `powershell -ExecutionPolicy Bypass -File scripts/build.ps1 -Release`。

5. **Excel / Copilot ブリッジ**
   - `scripts/excel_bridge.py` が ExcelManager 経由で実ワークブックを列挙します。Excel を起動しブックを開いてから `scripts/dev.ps1` を実行してください。
   - `scripts/copilot_bridge.py` は既存の `CopilotWorker` を利用して Playwright を起動します。初回実行時はサインイン後ブラウザを閉じずに残してください。
   - 既定で `python` コマンドを探します。異なるパスを使う場合は `SETTOUCHI_PYTHON=C:\\Python311\\python.exe` のように環境変数を設定します。

6. **実 Excel 連携を有効化するには**
   - `setouchi-excel` クレートに COM 実装を追加し、`MockExcelService` の代わりに実サービスを `AppWorker` へ注入してください。
