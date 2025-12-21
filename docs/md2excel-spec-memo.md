# Markdown→Excel 変換ツール 仕様メモ（md2excel）

## 1. ツール概要

- UTF-8 の Markdown ファイルを読み込み、
  **Excel 方眼紙風の 1 シート（"spec"）** に展開する Java ツール。
- 主な用途：
  - 仕様書っぽい Markdown を Excel 形式に整形する
  - 日本語前提、見出し／箇条書き／番号付き／テーブル／コードブロック／引用／水平線に対応
- Excel 書き込みは Apache POI（XSSF）。
- 入力は **List<String> の全読みではなく逐次処理**（メモリ使用量を抑える）。

---

## 2. 起動方法・入出力

### 2.1 起動パターン

1) **コマンドライン引数あり**

- `args[0]` : 入力 Markdown パス
- `args[1]` : 出力 xlsx パス（省略時は `.md` を `.xlsx` に置換）
- `args[2]` : 1行分として扱う列数（`mergeCols`、既定 40）
- `args[3]` : フォント名（省略時 `"游ゴシック"`）

※このモードではフォントサイズや縦位置は既定値を使用する。

2) **引数なし（jar ダブルクリック等）**

- Markdown ファイル選択（JFileChooser）
- 1行分の列数（`mergeCols`）を入力（既定 40）
- フォント選択ダイアログ
- セルの縦位置（上 / 中央 / 下）
- フォントサイズ入力
  - `#` 見出し（h1）
  - `##` 見出し（h2）
  - `###` 見出し（h3）
  - 通常テキスト（normal）

### 2.2 出力

- 出力先に `.xlsx` を生成。
- シート名は `"spec"` 固定。
- 処理完了後、絶対パス付きでダイアログ表示。

---

## 3. シート構成・レイアウト

- シート列数：`mergeCols`（既定 40 列）
- 各列幅：**3文字分**（`sheet.setColumnWidth(c, 3 * 256)`）
- グリッド線は非表示（画面＆印刷とも OFF）
- 各列のデフォルト列スタイルを `normalStyle` に設定：
  - `sheet.setDefaultColumnStyle(c, styles.normalStyle)`
- 「未使用セルを最後に埋める」処理は行わない（列デフォルトスタイルで見た目を統一する）。

---

## 4. 解析モデル（LineInfo）

各入力行は 1 回だけ解析し `LineInfo` を作成する（重複計算排除）。

### 4.1 LineInfo の内容

- `raw` : 元の行（インデント含む）
- `trimmed` : `raw.trim()`
- `indent` : 行頭のスペース/タブを数えた値（タブは 4 スペース換算）
- `kind` : 行種別（後述）
- 派生値（必要なものだけ保持）
  - 見出し：`headingLevel`, `headingText`
  - 引用：`quoteText`（先頭 `>` 除去後）
  - 箇条書き：`bulletMarkdownText`（`"・ "` 付与済み）

---

## 5. 行タイプの判定順（LineKind）

判定は `LineInfo.parse()` 内で次の順に行う。

1. **コードフェンス**：`trimmed.startsWith("```")`
2. **コードブロック中の行**：`st.inCodeBlock == true` → CODE_LINE
3. **空行**：`trimmed.isEmpty()`
4. **水平線**：`trimmed.equals("---")`
5. **引用**：`trimmed.startsWith(">")`
6. **テーブル行**：`MarkdownTable.isTableLine(raw)`
   - 区切り行：`MarkdownTable.isTableSeparatorLine(trimmed)` → TABLE_SEPARATOR
   - それ以外 → TABLE_ROW
7. **見出し**：`trimmed.startsWith("#")`
8. **箇条書き**：`trimmed.startsWith("* ")`
9. **番号付きリスト**：`MdTextUtil.isNumberedListLine(trimmed)`（正規表現なし）
10. 上記以外 → NORMAL

---

## 6. ブロック境界処理（MdBlockBoundary）

各行処理の先頭（render ループ側）で **必ず 1 回** `MdBlockBoundary.apply(li.kind.policy, ctx)` を呼び、
境界処理を統一する。handler 側では apply を呼ばない（＝呼び忘れを構造的に防ぐ）。

- `LineKind` は対応する `Policy` を保持する。
  - `LineKind` を追加する際は `Policy` 指定がコンパイル時に必須となり、apply 対応漏れを防げる。
- 実行順（order）は MdBlockBoundary の ORDER により固定。

### 6.1 境界 Action と順序

順序は以下で固定：

1. CLOSE_TABLE
2. CLOSE_BLOCK_QUOTE
3. INSERT_AUTO_BLANK_IF_PREV_HEADING
4. RESET_PARAGRAPH
5. CLEAR_LIST_CONTEXT

### 6.2 Policy と適用 Action

- CODE_FENCE：CLOSE_TABLE, CLOSE_BLOCK_QUOTE, RESET_PARAGRAPH
- MARKDOWN_BLANK：CLOSE_BLOCK_QUOTE, RESET_PARAGRAPH
- HORIZONTAL_RULE：CLOSE_TABLE, CLOSE_BLOCK_QUOTE, RESET_PARAGRAPH, CLEAR_LIST_CONTEXT
- HEADING：CLOSE_TABLE, CLOSE_BLOCK_QUOTE, RESET_PARAGRAPH, CLEAR_LIST_CONTEXT
- BULLET_ITEM：CLOSE_BLOCK_QUOTE, RESET_PARAGRAPH
- NUMBER_ITEM：CLOSE_BLOCK_QUOTE, INSERT_AUTO_BLANK_IF_PREV_HEADING, RESET_PARAGRAPH
- TABLE_LINE：CLOSE_BLOCK_QUOTE, RESET_PARAGRAPH
- （必要に応じて）NONE：何もしない

### 6.3 テーブル離脱検知

- 各行処理前に `closeTableIfLeaving(li.isTableLike(), ...)` を呼び、
  「テーブル行でなくなった瞬間」に `MarkdownTable.closeTableIfOpen` を呼ぶ。

---

## 7. 空行の扱い（RenderState に集約）

### 7.1 Markdown 空行（入力が空行）

`RenderState.onMarkdownBlankLine()` が唯一の入口。

- 直前が BLANK または HORIZONTAL_RULE の場合：
  - 行は増やさず Markdown 空行扱いとして状態だけ更新（従来仕様）
  - `afterConsumeMarkdownBlankWithoutNewRow()`
- それ以外：
  - 新規行を作成
  - `afterWriteMarkdownBlank(rowNum)`
- Markdown 由来の空行は `lastBlankFromMarkdown=true` かつ `lastBlankRowIndex` が reuse 対象になる。

### 7.2 自動空行（Markdown 由来ではない）

- 見出し前：`ensureAutoBlankBeforeHeadingIfNeeded()`
  - `rowIndex > 0 && lastRowType != BLANK` のときだけ 1 行挿入
- 見出し直後：`ensureAutoBlankIfPrevHeading()`
  - `lastRowType == HEADING` のときだけ 1 行挿入
- 自動空行は reuse 対象にしない：
  - `lastBlankFromMarkdown=false`, `lastBlankRowIndex=-1`

---

## 8. 要素別の描画仕様

> `<br>` の扱いは 8.7 を参照（見出し/箇条書き/番号/引用/テーブル/通常で統一）。

### 8.1 見出し（#〜）

- 書き込み列：常に A 列（col=0）
- 見出し前：必要なら自動空行 1 行（RenderState 関数）
- スタイル：
  - level 1 → heading1Style
  - level 2 → heading2Style
  - level 3 → heading3Style
  - level 4+ → heading4Style
- セル内容：`headingText` を `MarkdownInline.setMarkdownRichTextCell` で描画
  - 見出し内の `**` / `` ` `` は “そのまま MarkdownInline で解釈” される（除去しない）

#### 8.1.1 見出し内の `<br>`

- `MarkdownInline.splitByBrPreserveFormatting()` を必ず通し、分割結果を **A列に縦展開**する。
  - 例：`# 見出し1<br>見出し2` → A1=「見出し1」、A2=「見出し2」（同一見出しスタイル）
- 行末が `<br>` の場合、**次入力行へ見出しブロックを継続**する（スタイル維持）。
  - 太字などの継続は `carryPrefix` で保持する。

### 8.2 通常テキスト

- 見出し直後：必要なら自動空行 1 行（RenderState 関数）
- 列決定：
  - 見出し本文扱い（`inHeadingParagraphBlock && indent==0 && !inListBlock`）→ A 列
  - リスト注釈行（後述）→ A 列
  - リスト子段落（後述）→ C 列起点（2 + parentDepth）
  - それ以外：
    - indent==0 → A 列
    - listStack がある → `1 + depthForIndent`
    - listStack が空 → `1 + (indent/2)`
  - 最終的に `mergeLastCol-1` に clamp
- 追記ロジック（CellAppendUtil で追記）：
  1) 箇条書き説明行
  2) 番号付き説明行
  3) 同インデントの通常段落連結
  の順で判定し、該当すれば “同一セルへ追記” して行を作らない。

#### 8.2.1 通常テキスト内の `<br>`（説明行含む）

- `MarkdownInline.splitByBrPreserveFormatting()` を必ず通す。
- 1行目：既存セルへ追記 or 新規セルに出力（文脈に応じて）
- 2行目以降：**次行の同一列**に縦展開する。
- 行末が `<br>` の場合、**次入力行へ同一列で継続**する（太字などは carry で保持）。

### 8.3 箇条書き（`* `）

- `ListStackUtil.updateListDepth(listStack, indent, ordered=false)`
- 書き込み列：`1 + depth`（B 列開始）
- セル内容：`"・ " + content`（bulletMarkdownText）
- スタイル：`bulletStyle`

#### 8.3.1 箇条書き内の `<br>`

- `MarkdownInline.splitByBrPreserveFormatting()` を必ず通す（太字/インラインコードの扱いは 12章準拠）。
- 1行目：通常の箇条書きセル（B列〜）
- 2行目以降：**次行 + 1つ右の列**（C列〜）に縦展開する。
- 行末が `<br>` の場合、**次入力行（NORMAL 等）も右セル側へ継続**して吸う。
  - このモードでは既存の「箇条書き説明行の同一セル追記」は無効化する（暴発防止）。

### 8.4 番号付き（`1. ...`）

- 判定：`MdTextUtil.isNumberedListLine(trimmed)`（正規表現なし）
- `ListStackUtil.updateListDepth(listStack, indent, ordered=true)`
- 書き込み列：`1 + depth`（B 列開始）
- スタイル：`listStyle`
- 見出し直後の場合：
  - MdBlockBoundary.Policy.NUMBER_ITEM により “見出し直後の自動空行” が適用される

#### 8.4.1 番号付き内の `<br>`

- `MarkdownInline.splitByBrPreserveFormatting()` を必ず通す。
- 1行目：番号付きセル（B列〜）
- 2行目以降：**次行 + 1つ右の列**（C列〜）に縦展開する。
- 行末が `<br>` の場合、**次入力行（NORMAL 等）も右セル側へ継続**して吸う。
  - このモードでは既存の番号付き「説明行」追記（inNestedNumberBlock）は無効化する。

### 8.5 リストブロック後の注釈行／子段落

通常テキスト判定で `NormalTextFlags` を作り列決定・空行 reuse に使う。

- 注釈行（isListNote）：
  - 条件：`inListBlock && indent==0 && lastRowType==BLANK && lastBlankFromMarkdown && lastBlankRowIndex>=0`
  - 出力：A 列（0）
  - 空行は再利用する（reuseLastMarkdownBlankRow）
  - 出力後：`inListBlock=false`
- リスト子段落（isListChildParagraph）：
  - 条件：`indent>0 && inListBlock && lastRowType==BLANK && lastBlankFromMarkdown && lastBlankRowIndex>=0`
  - 出力列：C 列起点（2 + parentDepth）
  - 空行は再利用する

### 8.6 引用（`>`）

- `>` 先頭の行は引用扱い（テーブル判定より優先）
- 連続引用で `<br>` が無い場合、かつ直前が OTHER で Markdown 空行が挟まっていない場合：
  - 同一セルへ追記（appendMarkdownWithSpace）
- それ以外：
  - 列：`1 + depthForIndent(listStack, indent)`
  - セルに `quoteText` を出力
- ブロッククローズ：
  - ブロック終了時（境界処理など）に BlockQuoteUtil.closeBlockQuoteIfOpen
  - 引用範囲（firstRow..lastRow, startCol..lastColIndex）にスタイル適用：
    - 左端列だけ `blockQuoteLeftStyle`
    - それ以外は `blockQuoteBodyStyle`

#### 8.6.1 引用内の `<br>`

- `MarkdownInline.splitByBrPreserveFormatting()` を必ず通し、分割結果を **同一列に縦展開**する。
- 行末が `<br>` の場合、**次入力行（`>` 行でも通常行でも）を同一列に継続して吸う**。

### 8.7 `<br>` の共通ルール（見出し/箇条書き/番号/引用/テーブル/通常）

- `<br>` の検出・分割は **`MarkdownInline.splitByBrPreserveFormatting()` を唯一の入口**として統一する。
- `<br>` の判定実装は **MdTextUtil 側に集約**し、大小文字も含めて扱う（例：`<br>`, `<br/>`, `<br />`, `<BR>`, `<BR/>`, …）。
- **インラインコード（バッククォート内）に含まれる `<br>` はリテラル扱い**（分割しない）。
- テーブルは「分割して縦展開」ではなく、**セル内で連結（置換）**する（11章参照）。

---

## 9. コードブロック（```）

### 9.1 フェンス行

- ` ``` ` 行自体は出力しない（境界のみ）
- 開始側で `currentCodeBlockIndent = li.indent`
- `st.inCodeBlock` をトグルする

### 9.2 コード行（CODE_LINE）

- 書き込み列：
  - `depth = getDepthForIndent(listStack, currentCodeBlockIndent)`
  - `codeCol = clamp(1 + depth)`
- 左端トリム：
  - ブロック内最初の行の leadingSpaces を `codeBlockBaseIndent` として保持
  - `trimSpaces = min(leadingSpaces, codeBlockBaseIndent)`
  - `raw.substring(trimSpaces)` を表示
- スタイル：
  - `styles.codeBlockStyle`（背景グレー）
  - 文字ごとに ASCII/非ASCII 判定し Consolas / Meiryo を適用
- ブロック終了時：
  - `codeBlockFirstRow..LastRow` × `codeBlockCol..fillEndCol` に枠線スタイルを適用（mask方式）
- **コードブロック内の `<br>` は無視（置換/分割しない）**：
  - 入力の `<br>` はそのまま文字として表示する。

---

## 10. 水平線（---）

- 1行分を `horizontalRuleStyle` で埋める（指定列数）。
- 直前が水平線の空行は「行を増やさない」（7章）。

---

## 11. テーブル（|...|）

### 11.1 判定

- `MarkdownTable.isTableLine(raw)`：
  - `trim()` して `|` 始まり、かつ `|` が 2 つ以上
- 区切り行：
  - `MarkdownTable.isTableSeparatorLine(trimmed)` は “行を作らない”
  - ただし RenderState は「テーブル中」として更新する

### 11.2 配置列

- テーブル開始列（ヘッダ行だけ決定）：
  - `depth = getDepthForIndent(listStack, indent)`
  - `startCol = 1 + depth`
- 2行目以降：`currentTableStartCol` を継続使用

### 11.3 セル内容・スタイル

- 列分割は `split("\\|")` ではなく **手書きパース**で行う（正規表現依存を避ける）。
- ヘッダ行（最初の TABLE_ROW）：
  - `splitByBrPreserveFormatting()` を通したうえで、分割結果を **半角スペースで連結**して 1セルに収める
  - その後 `stripInlineMarkdown()`（`**` と `` ` `` の単純除去）してセル文字列にする
  - `tableHeaderStyle`
- ボディ行：
  - `splitByBrPreserveFormatting()` を通したうえで、分割結果を **半角スペースで連結**して 1セルに収める
  - 連結後の Markdown を `MarkdownInline.setMarkdownRichTextCell` で解釈（太字などは維持）
  - `tableBodyStyle`
- テーブル終了時：
  - 最終ボディ行に `tableBodyLastRowStyle` を適用（下線なし）

---

## 12. インライン Markdown（共通）

`MarkdownInline` が `**bold**` と `` `code` `` を解釈し、
XSSFRichTextString のフォント run を構築する。

- `**`：
  - “本物の太字マーカー” 判定あり（VS Code 互換寄りのガード）
- `` ` ``：
  - inCode 状態をトグル
- コード（inCode）：
  - ASCII 文字：Consolas + 赤
  - 非 ASCII：Meiryo + 赤
  - 太字指定（**`code`**）または baseStyle が太字なら bold 版のコードフォントを使用
- **インラインコード内の `<br>` は無視（置換/分割しない）**：
  - `` `a<br>b` `` は `<br>` をそのまま文字として表示する。

### 12.1 `<br>` 分割 API（共通）

- `MarkdownInline.splitByBrPreserveFormatting(markdown)` が `<br>` 処理の唯一の入口。
- `<br>` 判定は MdTextUtil の **唯一の実装**を使う（大小文字対応、`<br/>` 等も許容）。
- 分割時は太字の開閉を破綻させないように、必要に応じて行末に `**` を補い、
  次行には `carryPrefix`（例：`"**"`）で再開できる形にする。

---

## 13. SharedStrings 対策（追記時のクローン）

- Apache POI の RichText が共有される可能性があるため、
  `appendMarkdownToCell` は必ず RichText を clone してから追記する。
- これにより「別セルが連動して書き換わる」問題を防ぐ。

---

## 14. 状態管理（RenderState）

- 状態遷移は RenderState 内の `apply(Tx, ...)` に集約する。
- 行書き込み後は必ず `afterWriteXxx()` を呼ぶ。
- 空行／自動空行も RenderState API 経由で扱う。
- `<br>` による「次入力行への継続」を扱うための pending 状態を保持する：
  - 見出し継続（同スタイルでA列へ）
  - リスト継続（次行+1列右へ）
  - 引用継続（同列へ）
  - 同一列継続（通常/説明行の縦展開）
  - 太字などの継続は `carryPrefix` を保持して引き継ぐ

---
