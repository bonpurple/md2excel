# Release Notes - v2.0.0

リリース日: 2026-02-28
タグ: `v2.0.0`

本リリースは **パッケージ構成の整理により破壊的変更（breaking change）を含む** ため、SemVer に従いメジャーバージョンを更新しています。

---

## Highlights

- **パッケージ構成を再編**し、役割ごとにソースを整理しました。
- **箇条書き `- ` に対応**しました（従来の `* ` と同等にレンダリング）。
- **設定は常にダイアログ（GUI）で行う方式に変更**しました（CLI 引数による設定を廃止）。

---

## Added

- 箇条書き `- `（先頭が `-` + 空白の行）のレンダリングに対応

---

## Changed (Breaking)

- パッケージ構成を整理
  - `md2excel.app`（起動）
  - `md2excel.config`（設定）
  - `md2excel.excel`（Excel/POI ユーティリティ・Style）
  - `md2excel.markdown`（Markdown 文字処理・リスト深さ等）
  - `md2excel.render`（レンダリング本体）

- エントリポイント（main）の完全修飾クラス名（FQCN）が変更
  - **旧:** `md2excel.MarkdownToExcel`
  - **新:** `md2excel.app.MarkdownToExcel`

- パッケージ分割に伴い、外部参照が必要なクラス／メンバの可視性を調整（`public` 化）
  - 例: `MdStyle` のスタイル群、`RenderContext`、`MarkdownRenderer.render` など

- 実行時の設定取得方法を変更（GUI 化）
  - **常にダイアログで入力・選択**する方式に統一
  - CLI 引数モード（`input.md [output.xlsx] [mergeCols] [fontName]`）は **廃止**（引数は設定に使用しません）

---

## Upgrade Guide

### 起動クラス（FQCN）の更新

既存の起動クラスが `md2excel.MarkdownToExcel` の場合、`md2excel.app.MarkdownToExcel` に更新してください。

```text
java ... md2excel.app.MarkdownToExcel
```

### 実行時設定（GUI）

実行すると次の順にダイアログが出ます。

1. Markdown ファイル選択（`JFileChooser`）
2. `mergeCols` 入力
3. フォント選択
4. 縦位置選択
5. `# / ## / ### / 通常` のフォントサイズ入力

---

## Notes

- 本リポジトリは Maven/Gradle を使用しない前提のため、依存 jar の配置は `lib/README.md` を参照してください。
- headless 環境では Swing ダイアログ呼び出しを無効化する等の調整が必要な場合があります。

---
