package md2excel;

import java.util.ArrayList;
import java.util.List;

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

    final int mergeLastCol;
    final int lastColIndex;

    int rowIndex = 0;

    // リスト
    final List<ListStackUtil.ListLevel> listStack = new ArrayList<>();

    // 行種別
    RowType lastRowType = RowType.NONE;

    boolean inCodeBlock = false;
    boolean lastLineWasTable = false;

    boolean lastBlankFromMarkdown = false;
    int lastBlankRowIndex = -1;
    boolean lastBlankAfterTable = false;

    // 番号付き説明行
    boolean inNestedNumberBlock = false;
    int nestedNumberCol = 1;
    int nestedNumberIndent = 0;

    // 直前コンテンツ
    ContentType lastContentType = ContentType.NONE;
    int lastContentCol = 0;
    boolean lastContentWasTable = false;

    // コードブロック
    int codeBlockBaseIndent = -1;
    int codeBlockFirstRow = -1;
    int codeBlockLastRow = -1;
    int codeBlockCol = 0;
    int currentCodeBlockIndent = 0;

    // テーブル範囲
    int currentTableStartCol = 0;
    int currentTableHeaderRow = -1;
    int currentTableBodyStartRow = -1;
    int currentTableLastBodyRow = -1;
    int currentTableEndCol = -1;

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

    boolean pendingHeadingBr = false;
    int pendingHeadingLevel = 1;
    String pendingHeadingCarry = "";

    boolean pendingListBr = false;
    int pendingListBrCol = 0;
    int pendingListBrRow = -1;
    CellStyle pendingListBrStyle = null;
    String pendingListBrCarry = "";
    boolean pendingListBrHasCell = false;

    boolean pendingQuoteBr = false;
    int pendingQuoteBrCol = 0;
    String pendingQuoteBrCarry = "";

    boolean pendingSameColBr = false; // 説明行/通常文の <br> 継続
    int pendingSameColBrCol = 0;
    CellStyle pendingSameColBrStyle = null;
    String pendingSameColBrCarry = "";

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
        APPEND_BLOCKQUOTE_LINE,

        APPEND_TO_OPEN_BLOCKQUOTE_FROM_NORMAL,
        APPEND_NORMAL_TO_EXISTING_CELL,

        WRITE_NORMAL_TEXT
    }

    RenderState(int mergeCols) {
        this.mergeLastCol = mergeCols;
        this.lastColIndex = mergeCols - 1;
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
            lastContentCol = 0;
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

        case APPEND_BLOCKQUOTE_LINE:
            wroteOtherRow(false);

            lastContentType = ContentType.NORMAL;
            lastContentCol = blockQuoteCellCol;
            lastContentWasTable = false;

            lastWasBlockQuote = true;

            // 引用追記でも連結は切る
            cutParagraphLinking();
            return;

        case APPEND_TO_OPEN_BLOCKQUOTE_FROM_NORMAL:
            wroteOtherRow(false);

            lastContentType = ContentType.NORMAL;
            lastContentCol = col;
            lastContentWasTable = false;

            // “通常行→引用セル追記” は通常段落連結の起点（既存仕様維持）
            lastNormalRowIndex = rowNum;
            lastNormalIndent = 0;

            inHeadingParagraphBlock = false;

            // lastWasBlockQuote は触らない（既存仕様維持）
            return;

        case APPEND_NORMAL_TO_EXISTING_CELL:
            wroteOtherRow(false);

            lastContentType = ContentType.NORMAL;
            lastContentCol = col;
            lastContentWasTable = false;

            // indent==0 のときだけ「通常連結の起点」＋ bulletDetail 終了（既存仕様）
            if (indent == 0) {
                lastNormalRowIndex = rowNum;
                lastNormalIndent = 0;
                bulletDetailActive = false;
            }
            // lastWasBlockQuote は（従来も）ここでは触らない
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

        pendingHeadingBr = false;
        pendingHeadingCarry = "";

        pendingListBr = false;
        pendingListBrHasCell = false;
        pendingListBrRow = -1;
        pendingListBrCol = 0;
        pendingListBrStyle = null;
        pendingListBrCarry = "";

        pendingQuoteBr = false;
        pendingQuoteBrCol = 0;
        pendingQuoteBrCarry = "";

        pendingSameColBr = false;
        pendingSameColBrCol = 0;
        pendingSameColBrStyle = null;
        pendingSameColBrCarry = "";
    }

    void clearListContext() {
        inListBlock = false;

        inNestedNumberBlock = false;
        nestedNumberIndent = 0;
        nestedNumberCol = 1;

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

    int computeCodeTrimSpaces(int leadingSpaces) {
        if (codeBlockBaseIndent < 0)
            codeBlockBaseIndent = leadingSpaces;
        return Math.min(leadingSpaces, codeBlockBaseIndent);
    }

    void recordCodeBlockLinePos(int rowNum, int col) {
        if (codeBlockFirstRow < 0) {
            codeBlockFirstRow = rowNum;
            codeBlockCol = col;
        }
        codeBlockLastRow = rowNum;
    }

    void afterWriteBulletItem(int rowNum, int col) {
        apply(Tx.WRITE_BULLET_ITEM, rowNum, col, 0, false);
    }

    void afterWriteNumberedItem(int indent, int col) {
        apply(Tx.WRITE_NUMBERED_ITEM, -1, col, indent, false);
    }

    void afterAppendBlockQuoteLine() {
        apply(Tx.APPEND_BLOCKQUOTE_LINE, -1, -1, 0, false);
    }

    void afterWriteBlockQuoteLine(int rowNum, int col) {
        apply(Tx.WRITE_BLOCKQUOTE_LINE, rowNum, col, 0, false);
    }

    void afterAppendToOpenBlockQuoteFromNormalText(int rowNum, int col) {
        apply(Tx.APPEND_TO_OPEN_BLOCKQUOTE_FROM_NORMAL, rowNum, col, 0, false);
    }

    void afterAppendNormalToExistingCell(int rowNum, int col, int indent) {
        apply(Tx.APPEND_NORMAL_TO_EXISTING_CELL, rowNum, col, indent, false);
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
        // 連続空行 or 直前が水平線なら「行は増やさない」(従来仕様)
        if (lastRowType == RowType.BLANK || lastRowType == RowType.HORIZONTAL_RULE) {
            afterConsumeMarkdownBlankWithoutNewRow();
            return;
        }

        Row row = RowUtil.createRow(sheet, this, normalRowStyle);
        afterWriteMarkdownBlank(row.getRowNum());
    }

    /** 見出し前の自動空行：必要なときだけ入れる（従来仕様） */
    void ensureAutoBlankBeforeHeadingIfNeeded(Sheet sheet, CellStyle normalRowStyle) {
        if (rowIndex > 0 && lastRowType != RowType.BLANK) {
            writeAutoBlank(sheet, normalRowStyle);
        }
    }

    /** 「直前が見出しなら空行を1つ入れる」仕様（番号付き/通常文の見出し直後などで共用） */
    void ensureAutoBlankIfPrevHeading(Sheet sheet, CellStyle normalRowStyle) {
        if (lastRowType == RowType.HEADING) {
            writeAutoBlank(sheet, normalRowStyle);
        }
    }

    /** 自動空行を必ず1行書く（Markdown由来ではない、reuse対象にしない） */
    private void writeAutoBlank(Sheet sheet, CellStyle normalRowStyle) {
        Row row = RowUtil.createRow(sheet, this, normalRowStyle);
        afterWriteAutoBlank(row.getRowNum());
    }
}
