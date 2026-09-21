# md2excel（Markdown → Excel変換）

UTF-8のMarkdownファイルを読み込み、Apache POIを使用してExcelファイル（`.xlsx`）へ整形出力するツールです。

仕様書や設計書のMarkdownを、見出し・段落・リスト・引用・テーブル・コードブロックなどの構造を保ちながらExcelへ変換する用途を想定しています。

本リポジトリは個人による開発であり、特定の組織・雇用主とは無関係です。

- 入力：Markdown（UTF-8）
- 出力：Excel（`.xlsx`）
- シート名：`spec`
- 出力開始位置：B2

---

## 主な機能

- ATX形式・Setext形式の見出し
- 通常段落と複数行にまたがるインライン書式
- 箇条書き・番号付きリストとネスト
- 引用とネストした引用
- pipe形式のテーブル
- fenced code block（バッククォート・チルダ）
- 水平線
- 太字・斜体・インラインコード
- `<br>`、行末2スペース、行末バックスラッシュによる改行
- Markdown構造に応じたExcel上のインデント、背景色、罫線
- コード内のASCII・日本語に応じたフォント切り替え
- テーブルセル内の改行を複数のExcel行へ展開

引用内でも、見出し・リスト・テーブル・コードブロック・水平線などを利用できます。

---

## 動作環境

- Java 8以上
- Apache POI（XSSF / `XSSFWorkbook`）
- Swing（`JFileChooser` / `JOptionPane`）

本ツールはファイル選択や設定、完了通知にSwingを使用します。  
そのため、GUIを利用できないheadless環境では、そのまま実行できません。

---

## セットアップ

本リポジトリはMavenまたはGradleを使用していません。

Apache POIなどの依存jarはGitリポジトリに含まれていないため、[`lib/README.md`](lib/README.md) の手順に従って配置し、IDEのBuild Pathなどへ追加してください。

主なソースフォルダは次のとおりです。

```text
src/   アプリケーション
test/  JUnit 4テスト
```

---

## 使い方

### 実行

依存ライブラリをクラスパスへ追加したうえで、次のmainクラスを実行します。

```text
md2excel.app.MarkdownToExcel
```

IDEから実行するか、環境に合わせてJavaコマンドから実行してください。

```text
java ... md2excel.app.MarkdownToExcel
```

コマンドライン引数による設定には対応していません。

### 実行時の設定

次の順にダイアログが表示されます。

1. Markdownファイルの選択
2. シート左端（A列）基準の列数
3. フォント
4. セルの縦位置
5. 見出し・通常テキストのフォントサイズ

主な既定値は次のとおりです。

| 項目 | 既定値 |
|---|---:|
| シート列数 | 40 |
| フォント | 游ゴシック |
| 縦位置 | 下揃え |
| `#` 見出し | 16 pt |
| `##` 見出し | 14 pt |
| `###` 見出し | 12 pt |
| 通常テキスト | 11 pt |

シート列数は3～16,384、フォントサイズは5～72 ptの範囲で指定できます。

### 出力先

入力ファイルと同じディレクトリへ、拡張子を `.xlsx` に変更して出力します。

```text
document.md
    ↓
document.xlsx
```

同名のExcelファイルが存在する場合は上書きされます。

---

## 入出力例

### 入力

````markdown
## Title

This is **bold** and *italic*.

- item 1
- item 2<br>detail

> ### Quoted heading
>
> Quoted paragraph
>
> > Nested quote

```text
public class Sample {
    // code
}
```

| Name | Value |
| --- | --- |
| item | `value` |
````

### 出力

- 見出しはサイズと太字を反映
- 太字・斜体・インラインコードはセル内のリッチテキストとして出力
- リストや引用の深さに応じて列位置を調整
- 引用には背景色と左罫線を設定
- コードブロックには背景色と外周枠を設定
- テーブルにはヘッダー・本文用の罫線を設定
- `<br>`は次のExcel行へ展開

出力はB2から開始します。必要に応じて、生成後にExcel上で列幅などを調整してください。

---

## 対応している主な記法

### ブロック要素

- ATX見出し：`#`～`######`
- Setext見出し：`=` / `-`
- 箇条書き：`*` / `-` / `+`
- 番号付きリスト：`1.` / `1)` など
- 引用：`>`
- ネストした引用
- 水平線：`---` / `***` / `___` など
- pipe形式のテーブル
- fenced code block：`` ``` `` / `~~~`

### インライン要素

- 太字：`**text**` / `__text__`
- 斜体：`*text*` / `_text_`
- インラインコード：`` `code` ``
- 複数長バッククォートによるインラインコード
- バックスラッシュエスケープ
- `<br>`による改行
- 行末2スペースまたは行末`\`による改行

---

## テーブル利用時の注意

テーブルは、ヘッダー行の次に区切り行がある場合に認識されます。

```markdown
| Name | Value |
| --- | --- |
| A | 100 |
```

先頭・末尾の`|`は省略できます。

```markdown
Name | Value
--- | ---
A | 100
```

セル内に`|`を表示する場合は、`\|`と記述してください。

```markdown
| Expression |
| --- |
| A \| B |
```

インラインコード内であっても、未エスケープの`|`はセル区切りとして扱います。

---

## 制限事項

現在、次の項目には対応していません。

- リンク
- 画像
- 打ち消し線
- インデント形式のコードブロック
- Markdown内のHTML全般
- MarkdownまたはGFMの全構文

コードブロックはfenced code blockのみ対応しています。

本ツールはCommonMarkやGFMの完全な実装を目的としたものではありません。入力内容によっては、一般的なMarkdownレンダラーと異なる結果になる場合があります。

---

## テスト

### 自動テスト

`test/`にJUnit 4の自動テストがあります。

Eclipseでは、プロジェクトへJUnit 4を追加したうえで、次のクラスをJUnit Testとして実行します。

```text
md2excel.AllTests
```

主な確認対象は次のとおりです。

- Markdown文字処理
- コードフェンスとインラインコード
- リスト深度
- テーブル判定
- ネストした引用
- 設定値の検証
- Excel上のセル位置・値・フォント・罫線
- EOFで終了するコードブロック

### 目視確認

[`docs/verify-paragraph-rendering.md`](docs/verify-paragraph-rendering.md) は、生成されたExcelのレイアウトや装飾を確認するための受け入れテスト用Markdownです。

自動テストでは確認しにくい背景色、余白、全体的な読みやすさなどを目視確認します。

---

## ライセンス

Copyright (c) 2025 bonpurple

Apache License 2.0  
詳細は[`LICENSE`](LICENSE)を参照してください。

## NOTICE / Third-party

以下を参照してください。

- [`NOTICE`](NOTICE)
- [`docs/third-party/THIRD-PARTY-LICENSES.md`](docs/third-party/THIRD-PARTY-LICENSES.md)
- [`docs/third-party/THIRD-PARTY-NOTICES.txt`](docs/third-party/THIRD-PARTY-NOTICES.txt)