package md2excel.render;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import md2excel.markdown.ListStackUtil;

final class RenderState {

    enum RowType {
        NONE,
        BLANK,
        HEADING,
        HORIZONTAL_RULE,
        OTHER
    }

    enum ContentType {
        NONE,
        NORMAL,
        BULLET,
        NUMBER,
        CODE,
        HEADING,
        OTHER
    }

    final int renderEndColExclusive;
    final int renderLastColIndex;

    final int startRowIndex;
    final int startColIndex;

    int rowIndex;

    // リスト
    final List<ListStackUtil.ListLevel> listStack = new ArrayList<>();

    // 行種別
    RowType lastRowType = RowType.NONE;

    boolean lastLineWasTable = false;

    boolean lastBlankFromMarkdown = false;
    int lastBlankRowIndex = -1;
    boolean lastBlankAfterTable = false;

    private final CodeBlockState codeBlock = new CodeBlockState();
    private final TableState table = new TableState();

    CodeBlockState codeBlock() {
        return codeBlock;
    }

    TableState table() {
        return table;
    }

    // 番号付き説明行
    boolean inNestedNumberBlock = false;
    int nestedNumberCol;
    int nestedNumberIndent = 0;

    // 直前コンテンツ
    ContentType lastContentType = ContentType.NONE;
    int lastContentCol = 0;
    boolean lastContentWasTable = false;

    // 見出し本文
    boolean inHeadingParagraphBlock = false;

    // リストブロック中か
    boolean inListBlock = false;

    // 直近通常テキスト連結
    int lastNormalRowIndex = -1;
    int lastNormalIndent = -1;

    // 引用ブロック
    boolean inBlockQuote = false;
    int blockQuoteFirstRow = -1;
    int blockQuoteLastRow = -1;
    int blockQuoteCol = 0;
    boolean lastWasBlockQuote = false;
    int blockQuoteCellRow = -1;
    int blockQuoteCellCol = -1;

    // 箇条書き説明行（同一セル追記）
    boolean bulletDetailActive = false;
    int bulletDetailRow = -1;
    int bulletDetailCol = -1;

    // =========================
    // 状態遷移をここ1か所に集約
    // =========================
    private enum Tx {
        WRITE_MARKDOWN_BLANK,
        CONSUME_MARKDOWN_BLANK_NO_ROW,
        WRITE_AUTO_BLANK,

        WRITE_HORIZONTAL_RULE,
        WRITE_HEADING,
        WRITE_TABLE_ROW,
        SKIP_TABLE_SEPARATOR,

        WRITE_CODE_LINE,
        WRITE_BULLET_ITEM,
        WRITE_NUMBERED_ITEM,

        WRITE_BLOCKQUOTE_LINE,

        WRITE_NORMAL_TEXT
    }

    enum QuoteRowKind {
        NORMAL,
        HEADING_1,
        HEADING_2,
        HEADING_3,
        HEADING_4,
        BLANK,
        HORIZONTAL_RULE,
        TABLE,
        CODE
    }

    static final class QuoteRowInfo {

        final QuoteRowKind kind;
        final int depth;
        final int contentCol;

        final int tableStartCol;
        final int tableEndCol;
        final MarkdownTable.TableRowStyleRole tableRowStyleRole;

        QuoteRowInfo(QuoteRowKind kind, int depth, int contentCol) {

            this(kind, depth, contentCol, -1, -1, null);
        }

        QuoteRowInfo(QuoteRowKind kind, int depth, int contentCol, int tableStartCol, int tableEndCol,
                MarkdownTable.TableRowStyleRole tableRowStyleRole) {

            this.kind = kind;
            this.depth = Math.max(1, depth);
            this.contentCol = contentCol;
            this.tableStartCol = tableStartCol;
            this.tableEndCol = tableEndCol;
            this.tableRowStyleRole = tableRowStyleRole;
        }

        QuoteRowInfo withTableRowStyleRole(MarkdownTable.TableRowStyleRole role) {

            return new QuoteRowInfo(kind, depth, contentCol, tableStartCol, tableEndCol, role);
        }

        boolean isTableContentColumn(int col) {
            return tableStartCol >= 0 && tableEndCol >= tableStartCol && col >= tableStartCol && col <= tableEndCol;
        }
    }

    final Map<Integer, QuoteRowInfo> blockQuoteRows = new HashMap<Integer, QuoteRowInfo>();

    RenderState(int mergeCols) {
        this(mergeCols, 0, 0);
    }

    RenderState(int sheetColumnCount, int startRowIndex, int startColIndex) {

        this.startRowIndex = Math.max(0, startRowIndex);
        this.startColIndex = Math.max(0, startColIndex);

        int endColExclusive = sheetColumnCount - this.startColIndex;

        if (endColExclusive <= this.startColIndex) {
            endColExclusive = this.startColIndex + 1;
        }

        this.renderEndColExclusive = endColExclusive;
        this.renderLastColIndex = endColExclusive - 1;

        this.rowIndex = this.startRowIndex;
        this.nestedNumberCol = this.startColIndex + 1;
    }

    // 共通（「何かを書いた後」）の固定化。※ lastWasBlockQuote は呼び出し側（Tx）で決める
    private void wroteOtherRow(boolean table) {
        lastRowType = RowType.OTHER;
        lastLineWasTable = table;
        lastBlankFromMarkdown = false;
        lastBlankRowIndex = -1;
        lastBlankAfterTable = false;
    }

    // 「段落連結/箇条書き説明連結」を切る（安全側）
    private void cutParagraphLinking() {
        bulletDetailActive = false;
        lastNormalRowIndex = -1;
        lastNormalIndent = -1;
    }

    // ここが唯一の「状態遷移ルール本体」
    private void apply(Tx tx, int rowNum, int col, int indent, boolean isListNote) {
        switch (tx) {
        case WRITE_MARKDOWN_BLANK:
            lastRowType = RowType.BLANK;
            lastLineWasTable = false;
            lastBlankFromMarkdown = true;
            lastBlankRowIndex = rowNum; // reuse 対象
            lastBlankAfterTable = lastContentWasTable;
            // blank は直近コンテンツを更新しない
            return;

        case CONSUME_MARKDOWN_BLANK_NO_ROW:
            lastBlankFromMarkdown = true;
            if (lastRowType == RowType.BLANK && rowIndex > 0) {
                lastBlankRowIndex = rowIndex - 1; // 従来仕様：直前BLANKだけ reuse 合わせ
            }
            lastBlankAfterTable = lastContentWasTable;
            return;

        case WRITE_AUTO_BLANK:
            lastRowType = RowType.BLANK;
            lastLineWasTable = false;
            lastBlankFromMarkdown = false; // 重要：reuse 対象にしない
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;
            lastWasBlockQuote = false;
            return;

        case WRITE_HORIZONTAL_RULE:
            lastRowType = RowType.HORIZONTAL_RULE;
            lastLineWasTable = false;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;
            lastWasBlockQuote = false;
            lastContentWasTable = false;
            return;

        case WRITE_HEADING:
            lastRowType = RowType.HEADING;
            lastLineWasTable = false;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;

            lastContentType = ContentType.HEADING;
            lastContentCol = startColIndex;
            lastContentWasTable = false;

            inHeadingParagraphBlock = true;

            // 見出しは「連結」を切る（安全側）
            cutParagraphLinking();

            lastWasBlockQuote = false;
            return;

        case SKIP_TABLE_SEPARATOR:
            // 「行は書かないが table 中扱い」
            lastRowType = RowType.OTHER;
            lastLineWasTable = true;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;
            lastWasBlockQuote = false;
            lastContentWasTable = true;
            return;

        case WRITE_TABLE_ROW:
            wroteOtherRow(true);
            lastContentType = ContentType.OTHER;
            lastContentCol = col;
            lastContentWasTable = true;
            lastWasBlockQuote = false;
            return;

        case WRITE_CODE_LINE:
            wroteOtherRow(false);
            lastContentType = ContentType.CODE;
            lastContentCol = col;
            lastContentWasTable = false;
            lastWasBlockQuote = false;
            // コード行は連結を切る
            cutParagraphLinking();
            return;

        case WRITE_BULLET_ITEM:
            wroteOtherRow(false);
            inNestedNumberBlock = false;

            lastContentType = ContentType.BULLET;
            lastContentCol = col;
            lastContentWasTable = false;

            bulletDetailActive = true;
            bulletDetailRow = rowNum;
            bulletDetailCol = col;

            inListBlock = true;

            // 箇条書き開始で通常連結は切る
            lastNormalRowIndex = -1;
            lastNormalIndent = -1;

            lastWasBlockQuote = false;
            return;

        case WRITE_NUMBERED_ITEM:
            wroteOtherRow(false);
            bulletDetailActive = false;

            nestedNumberIndent = indent;
            nestedNumberCol = col;
            inNestedNumberBlock = true;

            lastContentType = ContentType.NUMBER;
            lastContentCol = col;
            lastContentWasTable = false;

            inListBlock = true;

            lastNormalRowIndex = -1;
            lastNormalIndent = -1;

            lastWasBlockQuote = false;
            return;

        case WRITE_BLOCKQUOTE_LINE:
            wroteOtherRow(false);

            if (!inBlockQuote) {
                inBlockQuote = true;
                blockQuoteFirstRow = rowNum;
                blockQuoteCol = col;
            }
            blockQuoteLastRow = rowNum;

            blockQuoteCellRow = rowNum;
            blockQuoteCellCol = col;

            lastContentType = ContentType.NORMAL; // quote は NORMAL 扱い
            lastContentCol = col;
            lastContentWasTable = false;

            lastWasBlockQuote = true;

            // 引用が来たら連結は切る
            cutParagraphLinking();
            return;

        case WRITE_NORMAL_TEXT:
            wroteOtherRow(false);

            lastContentType = ContentType.NORMAL;
            lastContentCol = col;
            lastContentWasTable = false;

            lastNormalRowIndex = rowNum;
            lastNormalIndent = indent;

            lastWasBlockQuote = false;

            if (isListNote)
                inListBlock = false;
            if (indent == 0)
                bulletDetailActive = false;
            return;
        }
    }

    void resetOnBlockBoundary() {
        // 段落境界でリセットしたいもの
        bulletDetailActive = false;
        lastNormalRowIndex = -1;
        lastNormalIndent = -1;
        // 「見出し本文ブロック」は段落境界で切る
        inHeadingParagraphBlock = false;
    }

    void clearListContext() {
        inListBlock = false;

        inNestedNumberBlock = false;
        nestedNumberIndent = 0;
        nestedNumberCol = startColIndex + 1;

        // ※ listStack は “インデント深さ計算” に使っているのでここでは消さない
        // （見出しで listStack を消すと、見出し後のインデント列決定が崩れる可能性があるため）
    }

    void afterWriteMarkdownBlank(int blankRowNum) {
        apply(Tx.WRITE_MARKDOWN_BLANK, blankRowNum, -1, 0, false);
    }

    void afterWriteHorizontalRule() {
        apply(Tx.WRITE_HORIZONTAL_RULE, -1, -1, 0, false);
    }

    void afterWriteHeading() {
        apply(Tx.WRITE_HEADING, -1, -1, 0, false);
    }

    void afterWriteTableRow(int startCol) {
        apply(Tx.WRITE_TABLE_ROW, -1, startCol, 0, false);
    }

    void afterWriteCodeLine(int col) {
        apply(Tx.WRITE_CODE_LINE, -1, col, 0, false);
    }

    void recordBlockQuoteRow(int rowNum, int quoteStartCol, int contentCol, QuoteRowKind kind, int depth) {

        int quoteDecorCol = quoteStartCol - 1;

        if (quoteDecorCol < 0) {
            quoteDecorCol = 0;
        }

        if (quoteDecorCol >= renderEndColExclusive) {
            quoteDecorCol = renderEndColExclusive - 1;
        }

        if (!inBlockQuote || blockQuoteFirstRow < 0) {
            inBlockQuote = true;
            blockQuoteFirstRow = rowNum;
            blockQuoteCol = quoteDecorCol;
        }

        if (quoteDecorCol < blockQuoteCol) {
            blockQuoteCol = quoteDecorCol;
        }

        blockQuoteLastRow = rowNum;

        blockQuoteRows.put(Integer.valueOf(rowNum), new QuoteRowInfo(kind, depth, contentCol));

        if (contentCol >= 0) {
            blockQuoteCellRow = rowNum;
            blockQuoteCellCol = contentCol;
        } else {
            blockQuoteCellRow = -1;
            blockQuoteCellCol = -1;
        }

        lastWasBlockQuote = true;
    }

    void recordBlockQuoteRow(int rowNum, int quoteStartCol, int contentCol, QuoteRowKind kind) {

        recordBlockQuoteRow(rowNum, quoteStartCol, contentCol, kind, 1);
    }

    void recordBlockQuoteTableRow(int rowNum, int quoteStartCol, int tableStartCol, int tableEndCol, int quoteDepth,
            MarkdownTable.TableRowStyleRole tableRowStyleRole) {

        // 引用範囲などの共通状態を更新する。
        recordBlockQuoteRow(rowNum, quoteStartCol, tableStartCol, QuoteRowKind.TABLE, quoteDepth);

        // テーブル固有の意味情報を含むメタ情報へ置き換える。
        blockQuoteRows.put(Integer.valueOf(rowNum), new QuoteRowInfo(QuoteRowKind.TABLE, quoteDepth, tableStartCol,
                tableStartCol, tableEndCol, tableRowStyleRole));
    }

    void updateBlockQuoteTableRowStyleRole(int rowNum, MarkdownTable.TableRowStyleRole tableRowStyleRole) {

        QuoteRowInfo current = blockQuoteRows.get(Integer.valueOf(rowNum));

        if (current == null || current.kind != QuoteRowKind.TABLE) {
            return;
        }

        blockQuoteRows.put(Integer.valueOf(rowNum), current.withTableRowStyleRole(tableRowStyleRole));
    }

    void afterWriteBulletItem(int rowNum, int col) {
        apply(Tx.WRITE_BULLET_ITEM, rowNum, col, 0, false);
    }

    void afterWriteNumberedItem(int indent, int col) {
        apply(Tx.WRITE_NUMBERED_ITEM, -1, col, indent, false);
    }

    void afterWriteBlockQuoteLine(int rowNum, int col) {
        apply(Tx.WRITE_BLOCKQUOTE_LINE, rowNum, col, 0, false);
    }

    void afterWriteNormalText(int rowNum, int col, int indent, boolean isListNote) {
        apply(Tx.WRITE_NORMAL_TEXT, rowNum, col, indent, isListNote);
    }

    // 自動挿入の空行（Markdown 由来ではない）を書いた後
    void afterWriteAutoBlank(int rowNum) {
        apply(Tx.WRITE_AUTO_BLANK, rowNum, -1, 0, false);
    }

    // 連続空行など「行は増やさない」が Markdown 空行扱いになるケース
    void afterConsumeMarkdownBlankWithoutNewRow() {
        apply(Tx.CONSUME_MARKDOWN_BLANK_NO_ROW, -1, -1, 0, false);
    }

    // テーブルの区切り行（|---|---|）は「行を書かないが table 中扱い」にする
    void afterSkipTableSeparatorLine() {
        apply(Tx.SKIP_TABLE_SEPARATOR, -1, -1, 0, false);
    }

    /** Markdown空行（入力の空行）を処理する：必要なら行を作り、必要なら作らない。 */
    void onMarkdownBlankLine(Sheet sheet, CellStyle normalRowStyle) {
        // 連続空行 or 直前が水平線なら「行は増やさない」
        if (lastRowType == RowType.BLANK || lastRowType == RowType.HORIZONTAL_RULE) {
            afterConsumeMarkdownBlankWithoutNewRow();
            return;
        }

        Row row = RowUtil.createRow(sheet, this, normalRowStyle);
        afterWriteMarkdownBlank(row.getRowNum());
    }

    /** 見出し前の自動空行：必要なときだけ入れる（従来仕様） */
    void ensureAutoBlankBeforeHeadingIfNeeded(Sheet sheet, CellStyle normalRowStyle) {
        if (rowIndex > startRowIndex && lastRowType != RowType.BLANK) {
            writeAutoBlank(sheet, normalRowStyle);
        }
    }

    /** 「直前が見出しなら空行を1つ入れる」仕様（番号付き/通常文の見出し直後などで共用） */
    void ensureAutoBlankIfPrevHeading(Sheet sheet, CellStyle normalRowStyle) {
        if (lastRowType == RowType.HEADING) {
            writeAutoBlank(sheet, normalRowStyle);
        }
    }

    /** 「直前が引用なら空行を1つ入れる」仕様 */
    void ensureAutoBlankIfPrevBlockQuote(Sheet sheet, CellStyle normalRowStyle) {
        if (lastWasBlockQuote && lastRowType != RowType.BLANK) {
            writeAutoBlank(sheet, normalRowStyle);
        }
    }

    /** 「直前がコード行なら空行を1つ入れる」仕様 */
    void ensureAutoBlankIfPrevCodeBlock(Sheet sheet, CellStyle normalRowStyle) {
        if (lastContentType == ContentType.CODE && lastRowType != RowType.BLANK) {
            writeAutoBlank(sheet, normalRowStyle);
        }
    }

    void ensureAutoBlankBeforeBlockQuoteIfNeeded(Sheet sheet, CellStyle blankRowStyle) {
        boolean prevNeedsSeparator = lastContentType == ContentType.NORMAL || lastContentType == ContentType.BULLET
                || lastContentType == ContentType.NUMBER || lastContentType == ContentType.HEADING;

        if (!inBlockQuote && !lastWasBlockQuote && rowIndex > startRowIndex && lastRowType != RowType.BLANK
                && prevNeedsSeparator) {
            writeAutoBlank(sheet, blankRowStyle);
        }
    }

    /** 直前のネストしたリスト（またはその説明行）が終わり、浅い階層のリストへ戻るか。 */
    boolean shouldInsertAutoBlankBeforeChildList(int currentIndent) {
        if (lastRowType == RowType.BLANK) {
            return false;
        }

        if (listStack.isEmpty()) {
            return false;
        }

        int prevIndent = listStack.get(listStack.size() - 1).indent;
        if (currentIndent >= prevIndent) {
            return false;
        }

        boolean cameFromNestedListContent = lastContentType == ContentType.BULLET
                || lastContentType == ContentType.NUMBER || bulletDetailActive || inNestedNumberBlock;

        return cameFromNestedListContent;
    }

    /** 直前のネストしたリスト（またはその説明行）が終わり、浅い階層のリストへ戻る場合は自動空行を1行入れる。 */
    void ensureAutoBlankBeforeChildListIfNeeded(Sheet sheet, CellStyle blankRowStyle, int currentIndent) {
        if (shouldInsertAutoBlankBeforeChildList(currentIndent)) {
            writeAutoBlank(sheet, blankRowStyle);
        }
    }

    /** 自動空行を必ず1行書く（Markdown由来ではない、reuse対象にしない） */
    private void writeAutoBlank(Sheet sheet, CellStyle normalRowStyle) {
        Row row = RowUtil.createRow(sheet, this, normalRowStyle);
        afterWriteAutoBlank(row.getRowNum());
    }

    void afterWriteQuotedHeading(int col) {
        apply(Tx.WRITE_HEADING, -1, -1, 0, false);

        lastContentCol = col;

        // 通常見出し用の「直後の通常段落」状態を引用外へ漏らさない。
        inHeadingParagraphBlock = false;

        lastWasBlockQuote = true;
    }

    QuoteRowInfo getBlockQuoteRowInfo(int rowNum) {
        return blockQuoteRows.get(Integer.valueOf(rowNum));
    }

    boolean isBlockQuoteRowKind(int rowNum, QuoteRowKind kind) {

        QuoteRowInfo info = getBlockQuoteRowInfo(rowNum);

        return info != null && info.kind == kind;
    }

    void clearBlockQuoteRows() {
        blockQuoteRows.clear();
    }
}