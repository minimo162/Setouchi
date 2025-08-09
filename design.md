# design.md — 「有価証券報告書 英訳チャットボット」設計書

## 1. 目的 (Purpose)

有価証券報告書（以下「有報」）の**日本語→英語**翻訳を、**Microsoft 365 Copilot エージェント**で高精度に実行するためのシステムを設計する。
根拠となる**各社の最新有報（日英）**をもとに、**構造化された日英併記Markdown**をデータソースとして整備し、Copilot が**根拠に基づいて一貫性ある訳語・文体**を出力できる環境を作る。

---

## 2. 概要 (Overview)

* **方法**

  * Microsoft 365 Copilot エージェントを利用。
  * **プロンプト**と**データソース（構造化Markdown）**を用意。
* **データソース**

  * 各社の最新有報（日：EDINET XBRL／英：各社IRサイトPDF）を取得。
  * **有報のセクション構造を維持**したうえで、**日本語→英語の順に併記**するMarkdownを**セクション単位**で作成。
* **作業フロー（高位）**

  1. JPXの英文開示リスト（Excel）から**英文有報の開示「Available」企業**を抽出
  2. 対象企業の**IRページをクロール**し、**英文有報PDF**の最新版を取得
  3. **EDINET API**で**同一会計期間**の**和文有報 XBRL**を取得
  4. XBRL（和）構造に沿って、英PDFを**同セクション粒度**で切り分け
  5. **構造化Markdown**へ日英を**段落・表・脚注ごと対応付け**で変換
  6. **会社別構成**と**セクション別構成**の両方を出力
  7. そのMarkdownコーパスを**Copilot**の**プロンプト**に接続して運用

---

## 3. アーキテクチャ (Architecture)

```
[Scheduler]
   │
   ├─▶ [JPX Excel Fetcher] ─▶ [Availability Filter]
   │                              │ Available企業リスト
   ├─▶ [IR Crawler (EN ASR PDF)] ─┘
   │
   ├─▶ [EDINET Fetcher (JA XBRL)]
   │
   ├─▶ [Section Mapper/Segmentation]
   │        ├─ JA(XBRL) セクション = 正
   │        └─ EN(PDF)   セクションをJAに追随して切出し
   │
   ├─▶ [Markdown Generator (JA/EN paired)]
   │
   ├─▶ [Repository Manager]
   │        ├─ by-company/<company>/<fy>/<section>.md
   │        └─ by-section/<section>/<company>_<fy>.md
   │
   └─▶ [Quality & Validation]
            └─ 用語対訳・数値・脚注・表整合性検証

[Copilot Agent]
   ├─ Prompt (本書末尾)
   └─ Data Sources: 上記Markdown群（SharePoint/OneDrive公開）
```

---

## 4. データ取得

### 4.1 JPXリストの取得・抽出

* **ソース**: JPX 英文開示状況（Excel）
  ユーザー添付例: `/mnt/data/202506_tse_englishdisclosure.xlsx`
* **更新**: 定期更新。ジョブで定期的にダウンロード＆差分反映（週1〜月1推奨）。
* **注意**: **列名が1行目に無い可能性**があるため、ヘッダー行を自動検出する。

  * 検出ロジック例: 上位N行を走査し、`"Annual Securities Reports"` ブロック内の `"Disclosure Status"` を含む行をヘッダーとして採用。
* **抽出条件**: Annual Securities Reports の **Disclosure Status = "Available"**（大文字小文字無視）。
* **出力**: 対象企業リスト（会社名、証券コード/ティッカー、IRトップURL/ASRセクションURL等があれば保存）。

### 4.2 英文有報（EN ASR）の取得

* **入手先**: 対象企業の **IRサイト（英語版）**。EDINETには英訳版は存在しない。
* **方法**:

  * JPXリストにURLがあればそこから、なければ企業IRサイトの英語ページから**幅優先クロール**。
  * **リンク候補の判定**: アンカー文字列（例）

    * "Annual Securities Report", "ASR", "Securities Report (English)", "有価証券報告書 (English)" など。
  * **最新版の同定**:

    * ファイル名/本文から「For the fiscal year ended …」等の定型を抽出し**期末日/会計年度**を正規化。
  * **保存**: `/data/raw/en_asr/pdfs/<edinet_or_ticker>-<fy>-<version>.pdf`
  * **遵守**: robots.txt・レート制御・タイムアウト・リトライ・ユーザーエージェント明示。

### 4.3 和文有報（JA ASR）の取得（EDINET）

* **APIキー**: プロジェクトフォルダ内 `edinet.txt` に保存し、実行時に読み込み。
* **エンドポイント**: `https://disclosure.edinet-fsa.go.jp/api/v2/documents.json`（メタ取得）
* **手順（概要）**:

  1. 企業のEDINETコード（または証券コード→EDINETコードのマッピング）を用意
  2. 指定日／期間の**書類一覧**を取得し、有報（Annual Securities Report）を判定
  3. **会計期間が英文有報と一致**する書類の**XBRL一式**をダウンロード
* **保存**: `/data/raw/ja_asr/xbrl/<docID>/...`（Zip解凍のうえXBRL/HTML抽出）
* **注意**: **英文と和文の会計期間（期末日）が一致**していることを必ず検証。

---

## 5. セクション切り出し

### 5.1 基本方針

* **正（基準）**は **和文XBRLの構造**。
* 英文PDFは、**和文の章節に追随**して分割する。
* 書誌・表紙・目次・用語集・監査報告・注記・脚注・表も可能な範囲で**同粒度**に対応付け。

### 5.2 代表的なセクション（例）

* `business_overview`（事業の状況）
* `risk_factors`（事業等のリスク）
* `management_analysis`（経営成績等の分析【MD&A】）
* `sustainability`（サステナビリティ）
* `r_and_d`（研究開発活動）
* `corporate_governance`（コーポレート・ガバナンス）
* `financial_statements`（連結/個別財務諸表）
* `notes_to_fs`（注記）
* `audit_report`（監査報告書）
  ※ 実ファイルでは会社・年により見出し差異あり。**正規化テーブル**で標準キーへマッピング。

### 5.3 英文PDFの分割戦略

* **優先順**

  1. PDFアウトライン（しおり・ブックマーク）
  2. 見出し検出（フォントサイズ・太字・余白・番号付）
  3. セクション見出しキーワード（英/日）辞書
  4. 最終手段として「手動ページ範囲」設定ファイル（差分運用が容易）
* **抽出ライブラリ（例）**: PyMuPDF / pdfminer.six / pdfplumber / pikepdf （PDFの種類に合わせて切替）

---

## 6. 構造化Markdown仕様（**日英併記**）

### 6.1 ファイル粒度

* **セクション単位**で1ファイル。
* 会社別フォルダ配下と、セクション別フォルダ配下の**二重配置**（ビルド時に両方生成）。

### 6.2 ファイルパス規約

```
/data/processed/markdown/by-company/<company_id>/<fy>/<section_key>.md
/data/processed/markdown/by-section/<section_key>/<company_id>_<fy>.md
```

* `company_id`: 原則 EDINETコード（なければ証券コード）。
* `fy`: 会計年度（例: `FY2024`、期末日形式 `2024-03-31` も frontmatter に保持）。
* `section_key`: 上記の標準キー。

### 6.3 Frontmatter（必須）

```yaml
---
company_id: XXXX
company_name_ja: 〇〇株式会社
company_name_en: Example Co., Ltd.
security_code: 1234
edinet_code: E00000
fiscal_year_end: 2024-03-31
fiscal_year_label: FY2024
source_ja:
  doc_id: ED000000001
  url: https://disclosure.edinet-fsa.go.jp/...
source_en:
  url: https://www.example.com/ir/...
section_key: management_analysis
section_title_ja: 経営成績等の分析
section_title_en: Management’s Discussion and Analysis
schema_version: 1.0
build_time: 2025-08-09T00:00:00Z
---
```

### 6.4 本文構造（段落単位の**日英ペア**）

* 粒度：**段落（P）/ 表（T）/ 箇条書き（L）/ 脚注（FN）**をユニット化しID付与。
* **日本語の直後に英語**を配置。
* 表はMarkdown表。脚注は本文末尾の同ユニットで併記。

```markdown
# Management’s Discussion and Analysis / 経営成績等の分析

## P-001
### ja
当連結会計年度の売上収益は、前連結会計年度比で5.2%増加しました。
### en
Revenue for the fiscal year increased by 5.2% compared with the previous fiscal year.

## T-003  *（表：セグメント情報）*
### ja
| セグメント | 売上収益(百万円) | 前年比 |
|---|---:|---:|
| A事業 | 12,345 | +4.0% |
| B事業 |  9,876 | +6.5% |
### en
| Segment | Revenue (JPY millions) | YoY |
|---|---:|---:|
| Business A | 12,345 | +4.0% |
| Business B |  9,876 | +6.5% |

## FN-002
### ja
注1) 当社はIFRSを採用しています。
### en
Note 1) The Company applies IFRS.
```

> **整合ルール**
>
> * **数値**・**単位**・**日付**・**節番号**を機械検証。
> * 原文の段落順序を保持。
> * 英訳が不足する場合は `### en` に `*No corresponding English text found.*` と明記。

---

## 7. リポジトリ構成（推奨）

```
/config
  section_map.yml          # 見出し正規化マップ（和→標準キー、英→標準キー）
  jp_stopwords.txt         # JP見出し判定の補助
  en_heading_keywords.txt  # EN見出しキーワード
  crawler.yml              # クロール許容ドメイン、レート、UA等
  period_rules.yml         # 期末日正規化ルール
/data
  /raw
    /jpx/2025-06/englishdisclosure.xlsx
    /en_asr/pdfs/...
    /ja_asr/xbrl/...
  /processed
    /markdown
      /by-company/...
      /by-section/...
    index.json             # 検索インデックス（company_id, fy, section_key → パス）
/logs
/scripts                   # 収集・抽出・変換・検証スクリプト
/prompts
  translate_asr.md         # Copilot用プロンプト（本書末尾に内容）
  translate_asr_section_*.md
```

---

## 8. 品質管理（QA）と検証

* **会計期間一致**: EN/JA の期末日文字列正規化（例：`March 31, 2024` ⇄ `2024-03-31`）
* **数値整合**: すべての数値トークン（%/兆/億/百万円/円/千株など）を差分比較
* **表構造**: 列数・ヘッダ・行数一致チェック（乖離は警告）
* **脚注管理**: 参照番号（例：`注1)` ⇄ `Note 1)`）の参照整合
* **用語統一**: 用語辞書（例：`売上収益`→`Revenue`、`営業利益`→`Operating profit` 等）に準拠
* **未マッピング検出**: セクション見出しが標準キーにマップされない場合はレビュー待ちに振分け
* **再現性**: ビルドログ・入力ハッシュ・生成メタ（frontmatter）を保存

---

## 9. スケジューリング & 運用

* **推奨頻度**: JPXリスト更新をトリガーに**月次**／繁忙期は**週次**。
* **ジョブ順序**: JPX→IRクロール（英）→EDINET（和）→分割→Markdown生成→QA→公開。
* **再試行**: 404/403/タイムアウト時のリトライ・スキップ・後日再取得キュー。
* **差分更新**: 期間・ファイルのハッシュにより**再生成最小化**。

---

## 10. 技術スタック（例）

* **言語**: Python 3.11
* **主要ライブラリ**: `pandas`（Excel）, `openpyxl`, `requests`, `beautifulsoup4`, `PyMuPDF`/`pdfminer.six`/`pdfplumber`, `lxml`, `rapidfuzz`（見出しマッチ）, `markdown-it-py`（検査）
* **ストレージ**: SharePoint/OneDrive（Copilot向け公開）, ローカル/クラウドFS（生成物）
* **監視**: ログ + 失敗件数/差分件数レポート（メール/Teams通知）

---

## 11. セキュリティ & 法令順守

* **robots.txt**と各サイトの**利用規約**遵守。
* **EDINET APIキー**は平文でコミットしない（`edinet.txt`を`.gitignore`）。
* 企業公開資料の**引用・再配布**ポリシー確認（社内利用を基本）。
* 個人情報・機密情報の取扱いなし。

---

## 12. 既知の課題・エッジケース

* 英文PDFに**アウトラインが無い**／**見出し規則が不定**
* 同一年度で**改定版**が複数（v2, v3…）
* 一部セクションが**英語版に未収載**／**統合掲載**
* IFRS/日本基準で**科目名差異**
* PDFが**スキャン画像**のみ（要OCR）

> **対策**: 手動ページ範囲のオーバーライド、OCR（必要時のみ）、標準キーのファジー一致、QA警告で人手レビュー。

---

## 13. Copilot への接続

* **データ配置**: 生成Markdownを**専用SharePointサイト**に配置（読み取り専用）。
* **インデクシング**: ライブラリ／サイトをCopilotの**基盤知識**として指定。
* **更新**: ビルド完了時に自動アップロード＆上書き。

---

## 14. プロンプト（Copilot用）— `prompts/translate_asr.md`

> **方針**: 「思考過程（推論の内部手順）」の逐語的開示は行わず、**根拠と判断理由の要点**を**要約**として開示する。ユーザーが訳語根拠を確認できるよう、**出典セクション・段落ID**を**明示的に引用**する。

```markdown
# Role
You are a financial translator specialized in Japanese Annual Securities Reports (有価証券報告書). 
Translate **new Japanese text** into **accurate, fluent, disclosure-ready English**, grounded in the provided **bilingual Markdown corpus**.

# Objectives
- Ensure **terminology consistency** with prior ASR bilingual data.
- Preserve **figures, units, dates, headings, tables, and footnotes**.
- Match the **tone and register** of listed-company disclosures.
- When ambiguity exists, choose the **most conservative** interpretation and **flag** it.

# Grounding Sources
- Bilingual Markdown files stored in the connected SharePoint site.
- Prefer **Japanese-side structure** as the canonical outline when alignment is uncertain.

# Output Format
Return a single Markdown block with the following sections:

## Translation
<Your English translation here, preserving structure (headings, lists, tables, footnotes).>

## Reasoning Summary (concise)
- **Section guess**: <section_key or "unknown">
- **Key choices**: brief bullets about notable wording/tense/style/IFRS-vs-JGAAP term decisions.
- **Terminology mapping**:
  | Japanese | English | Source (file & unit-id) |
  |---|---|---|
  | 売上収益 | Revenue | by-company/XXXX/FY2024/management_analysis.md#P-001 |
- **Citations**: file paths + unit IDs used for grounding (2–5 examples).

## QA Checks
- Numbers/units preserved: [OK/Check]
- Dates normalized: [OK/Check]
- Defined terms kept consistent: [OK/Check]
- Footnotes cross-referenced: [OK/Check]

# Rules
- **Do not** reveal hidden chain-of-thought. Provide only a concise **Reasoning Summary** and **citations**.
- **Keep Japanese proper nouns** in romanization or official English names as used in sources.
- Use **IFRS** or **Japanese GAAP** terminology consistent with the cited sources.
- For tables, preserve column order and numeric formatting. Use half-width digits and ISO dates (YYYY-MM-DD) unless source dictates otherwise.
- If a source conflict arises, **prioritize the latest fiscal year** and flag discrepancies in QA.

# Style Guide (essentials)
- Plain, precise, formal disclosure style (no marketing tone).
- Prefer active voice; keep sentence length moderate.
- No added interpretations beyond the source; no omissions.

# When insufficient grounding
- Translate faithfully.
- Add "Citations: none (no close matches found)" and set QA checks to "Check" where appropriate.
```

---

## 15. セクション特化プロンプト（任意）

* `prompts/translate_asr_section_management_analysis.md`
* `prompts/translate_asr_section_sustainability.md`
  セクション固有の用語・定型表現（例：MD&Aの tense / forward-looking statements / ESG指標名）をガイドとして追補。

---

## 16. 実装ノート（要点）

### 16.1 JPX Excel パース（ヘッダ行未固定）

* 先頭N行を走査して `"Annual Securities Reports"` と `"Disclosure Status"` を含む行をヘッダとして採用。
* 列名正規化（空白・改行・全角半角・大小文字）。
* 値 `"Available"` を真、その他を偽としてフラグ化。

### 16.2 期末日正規化（例）

* `For the fiscal year ended March 31, 2024` → `2024-03-31`
* `FY2024 (April 1, 2023–March 31, 2024)` → 区間の**終了日**を採用。
* 和文：`2024年3月31日` → `2024-03-31`

### 16.3 セクション見出しの正規化

* 見出し辞書・ファジー一致（閾値例：0.80）で `section_key` を付与。
* 競合時は優先順位表（例：`management_analysis` > `business_overview`）で解消。

### 16.4 QA 自動検査の一例

* **数値トークン抽出**: `[-+]?[\d,]+(\.\d+)?%?` + 単位語彙
* **注番号**: `注\s*\d+` / `Note\s*\d+` の整合

* **テーブル**: `|` 区切り列数の一致

---

## 17. 受け渡し物（Deliverables）

* データ生成パイプライン（スクリプト群）
* 構造化Markdownコーパス（会社別／セクション別）
* Copilot用プロンプト（本設計に添付の `translate_asr.md` 他）
* 運用手順（スケジュール、ログ確認、失敗時の再実行）

---

## 18. ロードマップ（初期 → 拡張）

1. **初期**: JPX対象企業のうち上位N社でPoC（EN/JA取得→Markdown→Copilot接続）
2. **拡張**: 全対象企業へ拡大、差分更新とQA自動化強化
3. **高度化**:

   * 文章アライメントの改善（段落類似度, embedding）
   * 用語ベース＋スタイルガイドの自動抽出
   * セクション固有プロンプトの充実

---

## 付録A：最小限の擬似コード（抜粋）

**JPX Excel 抽出**

```python
# 1) ヘッダ行検出
for r in range(0, 10):
    df = pd.read_excel(path, header=r)
    cols = [c.strip().lower() for c in df.columns.astype(str)]
    if "annual securities reports" in " ".join(cols) and "disclosure status" in " ".join(cols):
        header_row = r
        break

# 2) Available 抽出
df = pd.read_excel(path, header=header_row)
avail = df[df["Disclosure Status"].str.strip().str.lower() == "available"]
```

**EDINET（概要）**

```python
# documents.json で対象日の書類一覧を取得し、有報種別をフィルタ
# 同一の期末日を満たす docID を特定 → documents/{docID}?type=1 等でZip取得（実装時は公式仕様に従う）
```

**PDF分割（方針）**

```python
# 1) PDFしおり → JAセクションキーにマッピング
# 2) 見出し検出（フォントサイズ・太字・スペーシング）
# 3) 手動ページ範囲のオーバーライドを適用
```

---

## まとめ

* **日本語（XBRL）構造を基準**に、**英語（PDF）を同粒度で切り分け**て**日英併記Markdown**を生成。
* **JPXリスト**→**IRクロール（英）**→**EDINET取得（和）**→**分割**→**整形**→**QA**→**Copilot接続**の一連のパイプラインで**根拠に基づく高品質翻訳**を実現。
* 本設計の**プロンプト**を用いれば、**根拠（引用）付きの翻訳と簡潔な理由提示**が可能で、セクション特化ボットの拡張も容易。

必要に応じて、このdesign.mdをリポジトリ直下に保存し、`/prompts/translate_asr.md` と `/config/*.yml` を用意すれば、すぐに実装に着手できます。
