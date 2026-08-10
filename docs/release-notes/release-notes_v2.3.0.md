# Release Notes - v2.3.0

## Overview

v2.3.0 では、Markdown のブロック構造とインライン書式の解析精度を改善し、Excel への描画処理を段落単位の解析へ統一しました。

特に、複数行にまたがるインライン書式、Setext 見出し、引用ブロック内の各種ブロック要素、ネストした引用、GFM テーブルの pipe 解釈、テーブルセル内 `<br>` の描画を改善しています。

また、引用ブロックの行メタ情報と状態管理を `RenderState` に集約し、今後の構文追加や保守を行いやすい構造へ整理しました。

## Added

- Setext 見出しを通常見出しとして描画
  - `=` による level 1 見出し
  - `-` による level 2 見出し
  - 複数行 paragraph からの Setext 見出し化
  - inline 書式を含む Setext 見出し

- 引用ブロック内の各種ブロック要素に対応
  - 見出し
  - 水平線
  - fenced code block
  - テーブル
  - ネストした引用ブロック

- ネストした引用ブロックの深さに応じた描画
  - 引用深度ごとに左青色罫線を追加
  - 本文列を引用深度に応じて右へ移動

- Markdown 描画確認用の検証ファイルを追加
  - 通常段落
  - 引用
  - 箇条書き
  - 番号付きリスト
  - テーブル
  - コードブロック
  - Setext 見出し
  - 引用内の各 block type
  - block 境界パターン

## Changed

- Markdown 描画を段落単位解析へ統一
  - `ParagraphBuffer` / `ParagraphUtil` を導入
  - 通常文、引用、箇条書き、番号付きリストの継続行を段落として集約
  - soft break / hard break / `<br>` を段落単位で処理
  - 行をまたぐ inline 書式を保持
  - 見出し、セル追記、テーブル描画を paragraph API ベースへ統一

- テーブルセル内 `<br>` の描画を改善
  - 1つの Markdown テーブル行を複数の Excel 行へ展開
  - セルごとの `<br>` 数の差を空セルで補完
  - `<br>` によって生成された同一 Markdown 行内の Excel 行間では下罫線を描画しない
  - Markdown 上の別行との境界では従来通り下罫線を描画

- GFM テーブルの pipe 解釈を改善
  - 未エスケープ `|` をセル区切りとして処理
  - inline code 内の未エスケープ `|` もセル区切りとして扱う
  - body cell 数を header cell 数に合わせて切り詰め・補完

- 引用ブロック内のテーブル描画を改善
  - 通常テーブルと同じ table renderer を使用
  - 引用領域へグレー背景を適用
  - 指定終了列まで引用背景を描画

- 引用ブロック内のコードブロック描画を改善
  - 引用 marker を除去した内容を code line として処理
  - inline / list / table parsing を行わずコード本文をそのまま描画
  - 引用装飾と code block frame style を両立

- 見出し内の inline italic 描画を改善
  - 見出しの base font が bold の場合、italic segment でも bold を維持

## Fixed

- 行単位解析により、複数行 paragraph 内の emphasis が正しく解釈されない問題を修正

- 見出し直後のテーブル前に必要な空行が挿入されない問題を修正

- 引用ブロック内テーブルが通常文字列として描画される問題を修正

- 引用ブロック内の見出しが通常フォントで描画される問題を修正

- 引用ブロック内の水平線が文字列 `---` として描画される問題を修正

- 引用ブロック内水平線の描画位置を修正
  - 引用装飾列には水平線を描画しない
  - 本文側の列から水平線を開始
  - 空行フォントサイズを維持

- 引用ブロック内の fenced code block がコードブロックとして認識されない問題を修正

- ネストした引用ブロックの引用 marker が本文へ残る問題を修正

- ネストした引用ブロックの本文列と引用罫線位置を修正

## Refactoring

- 引用ブロック内の構文解析を通常ブロック解析と共通化
  - 引用 marker を1段ずつ除去
  - 内側の内容を通常と同じ `LineInfo` classifier へ通す
  - nested block quote を再帰的に解析

- 引用ブロックの行メタ情報と状態管理を `RenderState` に集約
  - 引用行種別
  - 引用深度
  - content column
  - 引用行登録処理
  - 引用終了時の cleanup

- `MarkdownRenderer` と `ParagraphUtil` に重複していた引用行登録処理を整理

- 旧 BR split / carry ベースの継続処理と不要な互換 API を整理

## Compatibility

- v2.2.4 からの機能拡張・描画精度改善リリースです。
- 公開 API の破壊的変更を目的としたリリースではありません。
- Markdown の解釈が CommonMark / GFM に近づいたため、従来と描画結果が変わるケースがあります。
  - Setext 見出し
  - GFM テーブル内の未エスケープ pipe
  - 引用ブロック内の block element
  - ネストした引用ブロック

## Verification

`docs/verify-paragraph-rendering.md` を使用し、以下を含む描画パターンを確認しています。

- 通常 paragraph の soft break / hard break
- 複数行にまたがる inline emphasis
- 引用 paragraph
- 引用内 heading / horizontal rule / fenced code block / table
- nested block quote
- bullet / numbered list
- table cell 内 `<br>`
- GFM table pipe
- Setext heading
- code block
- block boundary