# Changelog

このプロジェクトの主な変更履歴をまとめます。  
形式は [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) を参考にし、バージョニングは [SemVer](https://semver.org/lang/ja/) に準拠します。

---

## [2.0.0] - 2026-02-28

### Added
- 箇条書き `- `（先頭が `-` + 空白の行）のレンダリングに対応（既存の `* ` と同等の扱い）

### Changed
- パッケージ構成を整理（例: `md2excel.app` / `md2excel.config` / `md2excel.excel` / `md2excel.markdown` / `md2excel.render`）
- エントリポイント（main）の完全修飾クラス名が変更
  - `md2excel.MarkdownToExcel` → `md2excel.app.MarkdownToExcel`
- パッケージ分割に伴い、外部参照が必要なクラス／メンバを `public` 化（例: `MdStyle` のスタイル、`RenderContext`、`MarkdownRenderer.render` など）
- 設定の読み込み方式を変更：**常にダイアログ（GUI）で設定**するように変更（CLI 引数は使用しない）

### Removed
- CLI 引数モード（`input.md [output.xlsx] [mergeCols] [fontName]`）による設定

### Notes
- 起動クラス名（FQCN）や import は新しいパッケージに合わせて更新してください
- 本バージョンでは、実行時に引数を渡しても設定には使用されません（常にダイアログが表示されます）

---

## [1.0.0] - 2025-12-24

### Added
- Markdown（UTF-8）を Excel（.xlsx）に変換する基本機能（出力シート: `spec`）
- 見出し `#` / `##` / `###` / `####+` のスタイル出力
- 箇条書き `* ` と番号付きリスト `1. ` の出力（インデントに応じたネスト）
- 引用 `>` のブロック出力（左罫線＋背景）
- コードブロック（``` フェンス）の出力（背景＋外枠罫線、ASCII/CJK フォント切替）
- テーブル（`|...|`）の出力（ヘッダ/ボディの罫線・スタイル）
  - セル内 `\|` の復元、`` `...` `` 内の `|` を区切りにしない処理
- インライン装飾（セル内リッチテキスト）
  - 太字 `**...**`
  - インラインコード `` `...` ``（赤字、等幅、ASCII/CJK フォント切替）
- `<br>` の扱い
  - インラインコード外の `<br>` を改行として解釈し、縦方向展開
  - 行末 `<br>` による次行への継続（見出し／引用／リスト／通常文）
- 実行モード
  - CLI 引数モード: `input.md [output.xlsx] [mergeCols] [fontName]`
  - GUI 設定モード（引数なし）: ファイル選択、列数、フォント、縦位置、フォントサイズ設定

### Notes
- GUI（Swing）を使用するため、headless 環境では調整が必要な場合があります（完了ダイアログを含む）
- 箇条書きは `* ` のみ対応（`- ` / `+ ` は非対応）
- インライン Markdown は太字とインラインコードのみ対応（斜体・リンク等は未対応）
- コードブロックは ``` フェンスのみ対応

---
