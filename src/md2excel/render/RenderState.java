package md2excel.render;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

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

    private final RenderPositionState position;

    private final ListRenderState listState = new ListRenderState();

    private final BlockQuoteState blockQuoteState = new BlockQuoteState();

    private final CodeBlockState codeBlock = new CodeBlockState();

    private final TableState table = new TableState();

    // 行種別
    private RowType lastRowType = RowType.NONE;

    private boolean lastLineWasTable;

    private boolean lastBlankFromMarkdown;
    private int lastBlankRowIndex = -1;
    private boolean lastBlankAfterTable;

    // 直前コンテンツ
    private ContentType lastContentType = ContentType.NONE;

    private boolean lastContentWasTable;

    // 見出し本文
    private boolean inHeadingParagraphBlock;

    CodeBlockState codeBlock() {
        return codeBlock;
    }

    TableState table() {
        return table;
    }

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

        private final QuoteRowKind kind;
        private final int depth;
        private final int contentCol;

        private final int tableStartCol;
        private final int tableEndCol;
        private final MarkdownTable.TableRowStyleRole tableRowStyleRole;

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

        QuoteRowKind getKind() {
            return kind;
        }

        int getDepth() {
            return depth;
        }

        int getContentCol() {
            return contentCol;
        }

        MarkdownTable.TableRowStyleRole getTableRowStyleRole() {
            return tableRowStyleRole;
        }
    }

    RenderState(SheetColumnLayout columnLayout, int startRowIndex) {

        this.position = new RenderPositionState(columnLayout, startRowIndex);
    }

    int getRenderEndColExclusive() {
        return position.getRenderEndColExclusive();
    }

    int getRenderLastColIndex() {
        return position.getRenderLastColIndex();
    }

    int getStartColIndex() {
        return position.getStartColIndex();
    }

    int allocateNextRowIndex() {
        return position.allocateNextRowIndex();
    }

    int getPreviousRowIndex() {
        return position.getPreviousRowIndex();
    }

    boolean isLastRowType(RowType rowType) {
        return lastRowType == rowType;
    }

    boolean isLastLineTable() {
        return lastLineWasTable;
    }

    boolean isLastBlankAfterTable() {
        return lastBlankAfterTable;
    }

    boolean isLastContentType(ContentType contentType) {
        return lastContentType == contentType;
    }

    boolean isInHeadingParagraphBlock() {
        return inHeadingParagraphBlock;
    }

    boolean canReusePreviousQuotedBlank() {
        int previousRowIndex = position.getPreviousRowIndex();

        return lastRowType == RowType.BLANK && previousRowIndex >= 0
                && blockQuoteState.isRowKind(previousRowIndex, QuoteRowKind.BLANK);
    }

    void afterWriteQuotedBlank(int rowNum, int quoteStartCol, int contentCol, int quoteDepth) {

        recordBlockQuoteRow(rowNum, quoteStartCol, contentCol, QuoteRowKind.BLANK, quoteDepth);

        lastRowType = RowType.BLANK;
        lastLineWasTable = false;

        // 通常Markdown空行の再利用対象にはしない。
        lastBlankFromMarkdown = false;
        lastBlankRowIndex = -1;
        lastBlankAfterTable = false;

        lastContentType = ContentType.NORMAL;
        lastContentWasTable = false;

        listState.cutParagraphLinking();
        blockQuoteState.markLastWasBlockQuote();
    }

    void afterWriteQuotedHorizontalRule() {
        apply(Tx.WRITE_HORIZONTAL_RULE, -1, -1, 0, false);

        blockQuoteState.markLastWasBlockQuote();
    }

    void afterWriteQuotedCodeLine(int col) {
        apply(Tx.WRITE_CODE_LINE, -1, col, 0, false);

        blockQuoteState.markLastWasBlockQuote();
    }

    void afterWriteQuotedTableRow(int startCol) {
        afterWriteTableRow(startCol);
        blockQuoteState.markLastWasBlockQuote();
    }

    void afterSkipQuotedTableSeparatorLine() {
        apply(Tx.SKIP_TABLE_SEPARATOR, -1, -1, 0, false);

        blockQuoteState.markLastWasBlockQuote();
    }

    void afterWriteTableRow(int startCol) {
        apply(Tx.WRITE_TABLE_ROW, -1, startCol, 0, false);
    }

    void afterOpenNormalCodeFence() {
        lastLineWasTable = false;
    }

    void afterOpenQuotedCodeFence() {
        lastLineWasTable = false;
        blockQuoteState.markLastWasBlockQuote();
    }

    void afterFinishCodeBlock(boolean quoted) {
        lastLineWasTable = false;

        if (quoted) {
            blockQuoteState.markLastWasBlockQuote();
        }
    }

    void afterCloseTable() {
        lastLineWasTable = false;
        table.reset();
    }

    int updateListDepth(int indent, boolean ordered) {
        return listState.updateDepth(indent, ordered);
    }

    boolean hasListLevels() {
        return listState.hasLevels();
    }

    int getListDepthForIndent(int indent) {
        return listState.getDepthForIndent(indent);
    }

    int getParentListDepthForChildParagraph() {
        return listState.getParentDepthForChildParagraph();
    }

    boolean isInListBlock() {
        return listState.isInListBlock();
    }

    boolean wasLastBlockQuote() {
        return blockQuoteState.wasLastBlockQuote();
    }

    boolean hasRenderableBlockQuote() {
        return blockQuoteState.hasRenderableRows();
    }

    int getBlockQuoteFirstRow() {
        return blockQuoteState.getFirstRow();
    }

    int getBlockQuoteLastRow() {
        return blockQuoteState.getLastRow();
    }

    int getBlockQuoteStartCol() {
        return blockQuoteState.getStartCol();
    }

    void clearBlockQuoteTracking() {
        blockQuoteState.clearTracking();
    }

    boolean hasReusableMarkdownBlankForParagraph() {
        return lastRowType == RowType.BLANK && lastBlankFromMarkdown && lastBlankRowIndex >= 0;
    }

    boolean hasPreviousReusableMarkdownBlankRow() {
        return lastRowType == RowType.BLANK && lastBlankFromMarkdown && position.getNextRowIndex() > 0;
    }

    int getLastBlankRowIndex() {
        return lastBlankRowIndex;
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
        listState.cutParagraphLinking();
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
            if (lastRowType == RowType.BLANK && position.getNextRowIndex() > 0) {

                lastBlankRowIndex = position.getPreviousRowIndex();
            }
            lastBlankAfterTable = lastContentWasTable;
            return;

        case WRITE_AUTO_BLANK:
            lastRowType = RowType.BLANK;
            lastLineWasTable = false;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;

            blockQuoteState.markLastWasNotBlockQuote();
            return;

        case WRITE_HORIZONTAL_RULE:
            lastRowType = RowType.HORIZONTAL_RULE;
            lastLineWasTable = false;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;
            lastContentWasTable = false;

            blockQuoteState.markLastWasNotBlockQuote();
            return;

        case WRITE_HEADING:
            lastRowType = RowType.HEADING;
            lastLineWasTable = false;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;

            lastContentType = ContentType.HEADING;
            lastContentWasTable = false;

            inHeadingParagraphBlock = true;

            cutParagraphLinking();
            blockQuoteState.markLastWasNotBlockQuote();
            return;

        case SKIP_TABLE_SEPARATOR:
            lastRowType = RowType.OTHER;
            lastLineWasTable = true;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;
            lastContentWasTable = true;

            blockQuoteState.markLastWasNotBlockQuote();
            return;

        case WRITE_TABLE_ROW:
            wroteOtherRow(true);

            lastContentType = ContentType.OTHER;
            lastContentWasTable = true;

            blockQuoteState.markLastWasNotBlockQuote();
            return;

        case WRITE_CODE_LINE:
            wroteOtherRow(false);

            lastContentType = ContentType.CODE;
            lastContentWasTable = false;

            blockQuoteState.markLastWasNotBlockQuote();

            cutParagraphLinking();
            return;

        case WRITE_BULLET_ITEM:
            wroteOtherRow(false);

            lastContentType = ContentType.BULLET;
            lastContentWasTable = false;

            listState.afterWriteBulletItem();
            blockQuoteState.markLastWasNotBlockQuote();
            return;

        case WRITE_NUMBERED_ITEM:
            wroteOtherRow(false);

            lastContentType = ContentType.NUMBER;
            lastContentWasTable = false;

            listState.afterWriteNumberedItem();
            blockQuoteState.markLastWasNotBlockQuote();
            return;

        case WRITE_NORMAL_TEXT:
            wroteOtherRow(false);

            lastContentType = ContentType.NORMAL;
            lastContentWasTable = false;

            blockQuoteState.markLastWasNotBlockQuote();

            listState.afterWriteNormalText(isListNote, indent);

            return;
        }
    }

    void resetOnBlockBoundary() {
        listState.resetOnBlockBoundary();

        // 見出し本文ブロックは段落境界で切る。
        inHeadingParagraphBlock = false;
    }

    /**
     * 現在のリストブロックから離脱する。
     *
     * リストのインデント階層は、後続行の配置計算に使用するため保持する。
     */
    void leaveListBlockPreservingLevels() {
        listState.leaveBlockPreservingLevels();
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

    void afterWriteCodeLine(int col) {
        apply(Tx.WRITE_CODE_LINE, -1, col, 0, false);
    }

    void recordBlockQuoteRow(int rowNum, int quoteStartCol, int contentCol, QuoteRowKind kind, int depth) {

        int quoteDecorCol = quoteStartCol - 1;

        if (quoteDecorCol < 0) {
            quoteDecorCol = 0;
        }

        int renderEndColExclusive = position.getRenderEndColExclusive();

        if (quoteDecorCol >= renderEndColExclusive) {
            quoteDecorCol = renderEndColExclusive - 1;
        }

        blockQuoteState.recordRow(rowNum, quoteDecorCol, new QuoteRowInfo(kind, depth, contentCol));
    }

    void recordBlockQuoteRow(int rowNum, int quoteStartCol, int contentCol, QuoteRowKind kind) {

        recordBlockQuoteRow(rowNum, quoteStartCol, contentCol, kind, 1);
    }

    void recordBlockQuoteTableRow(int rowNum, int quoteStartCol, int tableStartCol, int tableEndCol, int quoteDepth,
            MarkdownTable.TableRowStyleRole tableRowStyleRole) {

        // 引用範囲などの共通状態を更新する。
        recordBlockQuoteRow(rowNum, quoteStartCol, tableStartCol, QuoteRowKind.TABLE, quoteDepth);

        // テーブル固有の意味情報を含むメタ情報へ置き換える。
        blockQuoteState.replaceRowInfo(rowNum, new QuoteRowInfo(QuoteRowKind.TABLE, quoteDepth, tableStartCol,
                tableStartCol, tableEndCol, tableRowStyleRole));
    }

    void updateBlockQuoteTableRowStyleRole(int rowNum, MarkdownTable.TableRowStyleRole tableRowStyleRole) {

        blockQuoteState.updateTableRowStyleRole(rowNum, tableRowStyleRole);
    }

    void afterWriteBulletItem(int rowNum, int col) {
        apply(Tx.WRITE_BULLET_ITEM, rowNum, col, 0, false);
    }

    void afterWriteNumberedItem(int indent, int col) {
        apply(Tx.WRITE_NUMBERED_ITEM, -1, col, indent, false);
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

    // 引用内の空行を、Excel行を追加せずに消費した場合
    void afterConsumeQuotedMarkdownBlankWithoutNewRow() {
        afterConsumeMarkdownBlankWithoutNewRow();
        blockQuoteState.markLastWasBlockQuote();
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
        if (position.hasWrittenRows() && lastRowType != RowType.BLANK) {
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

        if (blockQuoteState.wasLastBlockQuote() && lastRowType != RowType.BLANK) {

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

        if (!blockQuoteState.isOpen() && !blockQuoteState.wasLastBlockQuote() && position.hasWrittenRows()
                && lastRowType != RowType.BLANK && prevNeedsSeparator) {
            writeAutoBlank(sheet, blankRowStyle);
        }
    }

    /** 直前のネストしたリスト（またはその説明行）が終わり、浅い階層のリストへ戻るか。 */
    boolean shouldInsertAutoBlankBeforeChildList(int currentIndent) {

        if (lastRowType == RowType.BLANK) {
            return false;
        }

        boolean previousContentIsList = lastContentType == ContentType.BULLET || lastContentType == ContentType.NUMBER;

        return listState.shouldInsertAutoBlankBeforeChildList(currentIndent, previousContentIsList);
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

    void afterWriteQuotedHeading() {
        apply(Tx.WRITE_HEADING, -1, -1, 0, false);

        // 通常見出し用の状態を引用外へ漏らさない。
        inHeadingParagraphBlock = false;

        blockQuoteState.markLastWasBlockQuote();
    }

    QuoteRowInfo getBlockQuoteRowInfo(int rowNum) {
        return blockQuoteState.getRowInfo(rowNum);
    }

    boolean isBlockQuoteRowKind(int rowNum, QuoteRowKind kind) {

        return blockQuoteState.isRowKind(rowNum, kind);
    }
}