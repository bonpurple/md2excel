# md2excel

Markdown（UTF-8）をExcel（xlsx）へ変換するJavaツールです。見出し・箇条書き・番号付き・テーブル・コードブロック・引用・水平線に対応します。

## 使い方

### コマンドライン

```
java -jar md2excel.jar <input.md> [output.xlsx] [mergeCols] [fontName]
```

- `input.md`: 入力Markdown
- `output.xlsx`: 出力先（省略時は拡張子を`.xlsx`へ置換）
- `mergeCols`: 1行分の列数（既定40）
- `fontName`: フォント名（省略時は`游ゴシック`）

### 引数なし起動

jarをダブルクリック等で起動するとダイアログで設定できます。

## 出力

- シート名は`spec`固定
- グリッド線は非表示

## 開発メモ

詳細な仕様は`docs/md2excel-spec-memo.md`を参照してください。
