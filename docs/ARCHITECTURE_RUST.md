# Setouchi Copilot (Rust Edition)

## 技術スタック
- **Shell / Distribution**: Tauri 2 (Rust) + Tokio runtime
- **UI**: Leptos (CSR) + custom Setouchi Pearl componentキット
- **Excel連携**: `windows` crateでのCOM制御
- **Copilot自動化**: Playwright Rust bindings (Edge channel)
- **テレメトリ**: tracing + 日次ローテーションログ、`AppData/Roaming/Setouchi`

## ワークスペース構成
```
setouchi/
├─ apps/
│  └─ shell/          # Tauri + Leptos エントリポイント
├─ crates/
│  ├─ domain/         # DTO, 設定モデル, ドメインエラー
│  ├─ excel/          # COMラッパー, ワークブック制御
│  ├─ copilot/        # Playwrightを介したCopilot制御
│  └─ telemetry/      # ログ/設定保管ユーティリティ
└─ docs/              # 設計まとめ
```

## IPCフロー概要
1. UI から `submit_request` (tauri command) を `RequestMessage` で送信
2. `Tokio` channelで Excel/Copilot ワーカーに配送
3. ワーカーは `ResponseMessage` を broadcast channel 経由で UI へイベント送信
4. UI は `setouchi://response` イベントを listen し、Signal 更新/アニメーションを実行

## 今後の実装ステップ
1. Excel COM PoC: `GetObject("Excel.Application")` → ワークブック列挙/セル読み書き
2. Copilot自動化: Edge チャンネル、GPT-5トグル、チャット送信/レスポンス監視
3. Leptos UI: Hero/フォーム/タイムラインの再構築、状態管理/アニメーション導入
4. 自動テスト: `cargo test` + E2E (specta/Playwright) + 自動更新設定
