package md2excel.render;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

final class RenderState {

    enum RowType {
        NONE, BLANK, HEADING, HORIZONTAL_RULE, OTHER
    }

    enum ContentType {
        NONE, NORMAL, BULLET, NUMBER, CODE, HEADING, OTHER
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
        CONSUME_MARKDOWN_BLANK_NO_ROW, WRITE_AUTO_BLANK, WRITE_HORIZONTAL_RULE, WRITE_HEADING, WRITE_TABLE_ROW, SKIP_TABLE_SEPARATOR, WRITE_CODE_LINE, WRITE_BULLET_ITEM, WRITE_NUMBERED_ITEM
    }

    private enum RowOrigin {
        NORMAL, QUOTED
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

        return lastRowType == RowType.BLANK && previousRowIndex >= 0 && blockQuoteState.isRowKind(previousRowIndex, QuoteRowKind.BLANK);
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
        apply(Tx.WRITE_HORIZONTAL_RULE, RowOrigin.QUOTED);
    }

    void afterWriteQuotedCodeLine() {
        apply(Tx.WRITE_CODE_LINE, RowOrigin.QUOTED);
    }

    void afterWriteQuotedTableRow() {
        apply(Tx.WRITE_TABLE_ROW, RowOrigin.QUOTED);
    }

    void afterSkipQuotedTableSeparatorLine() {
        apply(Tx.SKIP_TABLE_SEPARATOR, RowOrigin.QUOTED);
    }

    void afterWriteTableRow() {
        apply(Tx.WRITE_TABLE_ROW);
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
    /**
     * 行番号が必要なMarkdown空行の遷移。
     *
     * 空行は直前のコンテンツ種別を更新せず、 直前が引用かどうかも変更しない。
     */
    private void applyMarkdownBlank(int rowNum) {
        lastRowType = RowType.BLANK;
        lastLineWasTable = false;
        lastBlankFromMarkdown = true;
        lastBlankRowIndex = rowNum;
        lastBlankAfterTable = lastContentWasTable;
    }

    /**
     * インデントとリスト注記の情報が必要な通常文の遷移。
     *
     * 引用段落の場合は、この後に引用行が記録され、 BlockQuoteState側で「直前は引用」に更新される。
     */
    private void applyNormalText(int indent, boolean isListNote) {
        wroteOtherRow(false);

        lastContentType = ContentType.NORMAL;
        lastContentWasTable = false;

        blockQuoteState.markLastWasNotBlockQuote();

        listState.afterWriteNormalText(isListNote, indent);
    }

    private void apply(Tx tx) {
        apply(tx, RowOrigin.NORMAL);
    }

    private void markLastOrigin(RowOrigin origin) {
        if (origin == RowOrigin.QUOTED) {
            blockQuoteState.markLastWasBlockQuote();
        } else {
            blockQuoteState.markLastWasNotBlockQuote();
        }
    }

    /**
     * 引数を必要としない遷移。
     *
     * Markdown空行と通常文は、それぞれ専用のメソッドで処理する。
     */
    private void apply(Tx tx, RowOrigin origin) {
        switch (tx) {
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

            // 自動空行は引用行ではない。
            blockQuoteState.markLastWasNotBlockQuote();
            return;

        case WRITE_HORIZONTAL_RULE:
            lastRowType = RowType.HORIZONTAL_RULE;
            lastLineWasTable = false;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;
            lastContentWasTable = false;

            markLastOrigin(origin);
            return;

        case WRITE_HEADING:
            lastRowType = RowType.HEADING;
            lastLineWasTable = false;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;

            lastContentType = ContentType.HEADING;
            lastContentWasTable = false;

            // 引用内見出しの状態を、引用外の「見出し本文」へ漏らさない。
            inHeadingParagraphBlock = origin == RowOrigin.NORMAL;

            cutParagraphLinking();
            markLastOrigin(origin);
            return;

        case SKIP_TABLE_SEPARATOR:
            lastRowType = RowType.OTHER;
            lastLineWasTable = true;
            lastBlankFromMarkdown = false;
            lastBlankRowIndex = -1;
            lastBlankAfterTable = false;
            lastContentWasTable = true;

            markLastOrigin(origin);
            return;

        case WRITE_TABLE_ROW:
            wroteOtherRow(true);

            lastContentType = ContentType.OTHER;
            lastContentWasTable = true;

            markLastOrigin(origin);
            return;

        case WRITE_CODE_LINE:
            wroteOtherRow(false);

            lastContentType = ContentType.CODE;
            lastContentWasTable = false;

            markLastOrigin(origin);
            cutParagraphLinking();
            return;

        case WRITE_BULLET_ITEM:
            wroteOtherRow(false);

            lastContentType = ContentType.BULLET;
            lastContentWasTable = false;

            listState.afterWriteBulletItem();
            markLastOrigin(origin);
            return;

        case WRITE_NUMBERED_ITEM:
            wroteOtherRow(false);

            lastContentType = ContentType.NUMBER;
            lastContentWasTable = false;

            listState.afterWriteNumberedItem();
            markLastOrigin(origin);
            return;

        default:
            throw new AssertionError("Unhandled transition: " + tx);
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
        applyMarkdownBlank(blankRowNum);
    }

    void afterWriteHorizontalRule() {
        apply(Tx.WRITE_HORIZONTAL_RULE);
    }

    void afterWriteHeading() {
        apply(Tx.WRITE_HEADING);
    }

    void afterWriteCodeLine() {
        apply(Tx.WRITE_CODE_LINE);
    }

    void recordBlockQuoteRow(int rowNum, int quoteStartCol, int contentCol, QuoteRowKind kind, int depth) {

        recordBlockQuoteRow(rowNum, quoteStartCol, new QuoteRowInfo(kind, depth, contentCol));
    }

    /**
     * 引用装飾列の計算と引用行の記録を共通化する。
     */
    private void recordBlockQuoteRow(int rowNum, int quoteStartCol, QuoteRowInfo rowInfo) {

        int quoteDecorCol = quoteStartCol - 1;

        if (quoteDecorCol < 0) {
            quoteDecorCol = 0;
        }

        int renderEndColExclusive = position.getRenderEndColExclusive();

        if (quoteDecorCol >= renderEndColExclusive) {
            quoteDecorCol = renderEndColExclusive - 1;
        }

        blockQuoteState.recordRow(rowNum, quoteDecorCol, rowInfo);
    }

    void recordBlockQuoteRow(int rowNum, int quoteStartCol, int contentCol, QuoteRowKind kind) {

        recordBlockQuoteRow(rowNum, quoteStartCol, contentCol, kind, 1);
    }

    void recordBlockQuoteTableRow(int rowNum, int quoteStartCol, int tableStartCol, int tableEndCol, int quoteDepth,
        TableRowStyleRole tableRowStyleRole) {

        QuoteRowInfo rowInfo = new QuoteRowInfo(QuoteRowKind.TABLE, quoteDepth, tableStartCol, tableStartCol, tableEndCol, tableRowStyleRole);

        recordBlockQuoteRow(rowNum, quoteStartCol, rowInfo);
    }

    void updateBlockQuoteTableRowStyleRole(int rowNum, TableRowStyleRole tableRowStyleRole) {

        blockQuoteState.updateTableRowStyleRole(rowNum, tableRowStyleRole);
    }

    void afterWriteBulletItem() {
        apply(Tx.WRITE_BULLET_ITEM);
    }

    void afterWriteNumberedItem() {
        apply(Tx.WRITE_NUMBERED_ITEM);
    }

    void afterWriteNormalText(int indent, boolean isListNote) {
        applyNormalText(indent, isListNote);
    }

    // 自動挿入の空行（Markdown 由来ではない）を書いた後
    void afterWriteAutoBlank() {
        apply(Tx.WRITE_AUTO_BLANK);
    }

    // 連続空行など「行は増やさない」が Markdown 空行扱いになるケース
    void afterConsumeMarkdownBlankWithoutNewRow() {
        apply(Tx.CONSUME_MARKDOWN_BLANK_NO_ROW);
    }

    // 引用内の空行を、Excel行を追加せずに消費した場合
    void afterConsumeQuotedMarkdownBlankWithoutNewRow() {
        afterConsumeMarkdownBlankWithoutNewRow();
        blockQuoteState.markLastWasBlockQuote();
    }

    // テーブルの区切り行（|---|---|）は「行を書かないが table 中扱い」にする
    void afterSkipTableSeparatorLine() {
        apply(Tx.SKIP_TABLE_SEPARATOR);
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
        boolean prevNeedsSeparator = lastContentType == ContentType.NORMAL || lastContentType == ContentType.BULLET || lastContentType == ContentType.NUMBER
            || lastContentType == ContentType.HEADING;

        if (!blockQuoteState.isOpen() && !blockQuoteState.wasLastBlockQuote() && position.hasWrittenRows() && lastRowType != RowType.BLANK
            && prevNeedsSeparator) {
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
        RowUtil.createRow(sheet, this, normalRowStyle);
        afterWriteAutoBlank();
    }

    void afterWriteQuotedHeading() {
        apply(Tx.WRITE_HEADING, RowOrigin.QUOTED);
    }

    QuoteRowInfo getBlockQuoteRowInfo(int rowNum) {
        return blockQuoteState.getRowInfo(rowNum);
    }

    boolean isBlockQuoteRowKind(int rowNum, QuoteRowKind kind) {

        return blockQuoteState.isRowKind(rowNum, kind);
    }
}
