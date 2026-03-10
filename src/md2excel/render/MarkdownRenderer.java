package md2excel.render;

import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import md2excel.excel.Md2ExcelSheetUtil;
import md2excel.markdown.ListStackUtil;
import md2excel.markdown.MdTextUtil;

public final class MarkdownRenderer {

    private enum LineKind {
        CODE_FENCE(MdBlockBoundary.Policy.CODE_FENCE),
        CODE_LINE(MdBlockBoundary.Policy.NONE), // inCodeBlock中は境界処理しない（従来通り）
        BLANK(MdBlockBoundary.Policy.MARKDOWN_BLANK),
        HORIZONTAL_RULE(MdBlockBoundary.Policy.HORIZONTAL_RULE),
        BLOCK_QUOTE(MdBlockBoundary.Policy.NONE), // 従来 apply していないなら NONE
        TABLE_SEPARATOR(MdBlockBoundary.Policy.TABLE_LINE),
        TABLE_ROW(MdBlockBoundary.Policy.TABLE_LINE),
        HEADING(MdBlockBoundary.Policy.HEADING),
        BULLET_ITEM(MdBlockBoundary.Policy.BULLET_ITEM),
        NUMBER_ITEM(MdBlockBoundary.Policy.NUMBER_ITEM),
        NORMAL(MdBlockBoundary.Policy.NONE); // 従来 apply していないなら NONE

        final MdBlockBoundary.Policy policy;

        LineKind(MdBlockBoundary.Policy policy) {
            this.policy = policy;
        }
    }

    private static final class LineInfo {
        final String raw; // 元行（インデント含む）
        final String trimmed; // raw.trim()
        final int indent; // leading spaces/tabs（raw基準）
        final LineKind kind;

        // kind別で使う派生値（必要なものだけ埋める）
        final int headingLevel; // kindがHEADINGのときのみ >=1
        final String headingText; // kindがHEADINGのときのみ
        final String quoteText; // kindがBLOCK_QUOTEのときのみ（">"除去済み）
        final String bulletMarkdownText;// kindがBULLET_ITEMのときのみ（"・ "付与済み）

        private LineInfo(String raw, String trimmed, int indent, LineKind kind, int headingLevel, String headingText,
                String quoteText, String bulletMarkdownText) {
            this.raw = raw;
            this.trimmed = trimmed;
            this.indent = indent;
            this.kind = kind;

            this.headingLevel = headingLevel;
            this.headingText = headingText;
            this.quoteText = quoteText;
            this.bulletMarkdownText = bulletMarkdownText;
        }

        boolean isTableLike() {
            return kind == LineKind.TABLE_SEPARATOR || kind == LineKind.TABLE_ROW;
        }

        static LineInfo parse(String rawLine, RenderState st) {
            String trimmed = rawLine.trim();
            int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);

            // 1) code fence は inCodeBlock 中でも最優先（閉じるため）
            if (trimmed.startsWith("```") || trimmed.startsWith("~~~")) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_FENCE, -1, null, null, null);
            }

            // 2) code block 中は「全部 code line」
            if (st.inCodeBlock) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_LINE, -1, null, null, null);
            }

            // 3) blank
            if (trimmed.isEmpty()) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.BLANK, -1, null, null, null);
            }

            // 4) horizontal rule
            if (MdTextUtil.isHorizontalRuleLine(trimmed)) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.HORIZONTAL_RULE, -1, null, null, null);
            }

            // 5) block quote（テーブル判定より先："> |a|b|" は引用扱い）
            if (trimmed.startsWith(">")) {
                String quoteText = trimmed.substring(1).trim();
                return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null, quoteText, null);
            }

            // 6) table（isTableLine をここで1回だけ）
            if (MarkdownTable.isTableLine(rawLine)) {
                boolean sep = MarkdownTable.isTableSeparatorLine(trimmed);
                return new LineInfo(rawLine, trimmed, indent, sep ? LineKind.TABLE_SEPARATOR : LineKind.TABLE_ROW, -1,
                        null, null, null);
            }

            // 7) heading
            if (trimmed.startsWith("#")) {
                int level = MdTextUtil.countHeadingLevel(trimmed);
                String text = trimmed.substring(level).trim();
                text = MdTextUtil.stripHeadingClosingHashes(text);
                return new LineInfo(rawLine, trimmed, indent, LineKind.HEADING, level, text, null, null);
            }

            // 8) list
            if (trimmed.length() >= 2) {
                char m = trimmed.charAt(0);
                if ((m == '*' || m == '-' || m == '+') && Character.isWhitespace(trimmed.charAt(1))) {
                    String content = trimmed.substring(2).trim();
                    String bulletMd = "・ " + content;
                    return new LineInfo(rawLine, trimmed, indent, LineKind.BULLET_ITEM, -1, null, null, bulletMd);
                }
            }

            if (MdTextUtil.isNumberedListLine(trimmed)) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.NUMBER_ITEM, -1, null, null, null);
            }

            return new LineInfo(rawLine, trimmed, indent, LineKind.NORMAL, -1, null, null, null);
        }
    }

    public static void render(Iterator<String> it, RenderContext ctx) {
        RenderState st = ctx.st;

        while (it.hasNext()) {
            String rawLine = it.next();
            LineInfo li = LineInfo.parse(rawLine, st);

            MdBlockBoundary.closeTableIfLeaving(li.isTableLike(), ctx);

            // ここで必ず境界処理を実施（呼び忘れが起きない）
            MdBlockBoundary.apply(li.kind.policy, ctx);

            if (tryConsumeHeadingBr(li, ctx))
                continue;
            if (tryConsumeListBr(li, ctx))
                continue;
            if (tryConsumeQuoteBr(li, ctx))
                continue;
            if (tryConsumeSameColBr(li, ctx))
                continue;

            switch (li.kind) {
            case CODE_FENCE:
                handleCodeFence(li, ctx);
                break;
            case CODE_LINE:
                handleInCodeBlock(li, ctx);
                break;
            case BLANK:
                handleBlankLine(li, ctx);
                break;
            case HORIZONTAL_RULE:
                handleHorizontalRule(li, ctx);
                break;
            case BLOCK_QUOTE:
                handleBlockQuote(li, ctx);
                break;
            case TABLE_SEPARATOR:
                handleTableSeparatorLine(li, ctx);
                break;
            case TABLE_ROW:
                handleTableRow(li, ctx);
                break;
            case HEADING:
                handleHeading(li, ctx);
                break;
            case BULLET_ITEM:
                handleBullet(li, ctx);
                break;
            case NUMBER_ITEM:
                handleNumberedList(li, ctx);
                break;
            case NORMAL:
                handleNormalText(li, ctx);
                break;
            default:
                throw new AssertionError("Unhandled LineKind: " + li.kind);
            }
        }

        if (st.lastLineWasTable) {
            MarkdownTable.closeTableIfOpen(ctx.sheet, ctx.styles, st);
        }
        BlockQuoteUtil.closeBlockQuoteIfOpen(ctx.sheet, ctx.styles, st);
    }

    private static void handleCodeFence(LineInfo li, RenderContext ctx) {

        if (!ctx.st.inCodeBlock) {
            ctx.st.ensureAutoBlankIfPrevBlockQuote(ctx.sheet, ctx.styles.normalStyle);
            ctx.st.currentCodeBlockIndent = li.indent;
        }

        if (ctx.st.inCodeBlock && ctx.st.codeBlockFirstRow >= 0 && ctx.st.codeBlockLastRow >= 0) {
            int fillEndCol = Math.max(ctx.st.codeBlockCol, ctx.st.lastColIndex);

            for (int r = ctx.st.codeBlockFirstRow; r <= ctx.st.codeBlockLastRow; r++) {
                Row rowObj = ctx.sheet.getRow(r);
                if (rowObj == null)
                    continue;

                for (int c = ctx.st.codeBlockCol; c <= fillEndCol; c++) {
                    Cell cell = rowObj.getCell(c);
                    if (cell == null)
                        cell = rowObj.createCell(c);

                    boolean isTop = (r == ctx.st.codeBlockFirstRow);
                    boolean isBottom = (r == ctx.st.codeBlockLastRow);
                    boolean isLeft = (c == ctx.st.codeBlockCol);
                    boolean isRight = (c == fillEndCol);

                    int mask = 0;
                    if (isTop)
                        mask |= 1;
                    if (isBottom)
                        mask |= 2;
                    if (isLeft)
                        mask |= 4;
                    if (isRight)
                        mask |= 8;

                    cell.setCellStyle(ctx.styles.codeBlockFrameStyle(mask));
                }
            }
        }

        ctx.st.inCodeBlock = !ctx.st.inCodeBlock;
        ctx.st.lastLineWasTable = false;

        ctx.st.codeBlockFirstRow = -1;
        ctx.st.codeBlockLastRow = -1;
        ctx.st.codeBlockCol = 0;
        ctx.st.codeBlockBaseIndent = -1;
    }

    private static void handleInCodeBlock(LineInfo li, RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.CODE_LINE,
                ctx.styles.normalStyle);

        int codeCol = calcBlockStartCol(ctx.st.currentCodeBlockIndent, ctx.st);

        int leadingSpaces = li.indent;
        int trimSpaces = ctx.st.computeCodeTrimSpaces(leadingSpaces);
        String codeLine = li.raw.substring(trimSpaces);

        Cell cell = row.createCell(codeCol);
        MarkdownInline.setCodeBlockRichTextCell(ctx.wb, cell, codeLine, ctx.styles.codeBlockStyle);

        ctx.st.recordCodeBlockLinePos(row.getRowNum(), codeCol);
        ctx.st.afterWriteCodeLine(codeCol);
    }

    private static void handleBlankLine(LineInfo li, RenderContext ctx) {
        ctx.st.onMarkdownBlankLine(ctx.sheet, ctx.styles.normalStyle);
    }

    private static void handleHorizontalRule(LineInfo li, RenderContext ctx) {
        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx, RowUtil.ReuseKind.HORIZONTAL_RULE,
                ctx.styles.normalStyle);
        Md2ExcelSheetUtil.createHorizontalRuleRow(ctx.sheet, row, ctx.styles.horizontalRuleStyle, ctx.st.mergeLastCol);
        ctx.st.afterWriteHorizontalRule();
    }

    private static void handleBlockQuote(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.normalStyle);

        String quoteText = applyHardLineBreak(li.quoteText, li);
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(quoteText);
        boolean hasBr = sp.endsWithBr || sp.lines.size() >= 2;

        if (!hasBr && ctx.st.inBlockQuote && ctx.st.blockQuoteCellRow >= 0 && ctx.st.blockQuoteCellCol >= 0
                && ctx.st.lastRowType == RenderState.RowType.OTHER && !ctx.st.lastBlankFromMarkdown) {

            appendBrSplitLineWithSpace(ctx, ctx.st.blockQuoteCellRow, ctx.st.blockQuoteCellCol, sp, 0,
                    ctx.styles.normalStyle);

            ctx.st.afterAppendBlockQuoteLine();
            return;
        }

        int col = calcBlockStartCol(li.indent, ctx.st);

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.BLOCK_QUOTE,
                ctx.styles.normalStyle);
        Cell cell = row.createCell(col);
        setBrSplitLineCell(ctx, cell, sp, 0, ctx.styles.normalStyle);
        ctx.st.afterWriteBlockQuoteLine(row.getRowNum(), col);

        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(col);
            setBrSplitLineCell(ctx, c2, sp, i, ctx.styles.normalStyle);
            ctx.st.afterWriteBlockQuoteLine(r2.getRowNum(), col);
        }

        ctx.st.pendingQuoteBr = sp.endsWithBr;
        ctx.st.pendingQuoteBrCol = col;
        ctx.st.pendingQuoteBrCarry = sp.carryPrefix;
    }

    private static boolean tryConsumeQuoteBr(LineInfo li, RenderContext ctx) {
        if (!ctx.st.pendingQuoteBr)
            return false;

        if (li.kind != LineKind.BLOCK_QUOTE && li.kind != LineKind.NORMAL) {
            ctx.st.pendingQuoteBr = false;
            ctx.st.pendingQuoteBrCarry = "";
            return false;
        }

        String text = (li.kind == LineKind.BLOCK_QUOTE) ? li.quoteText : li.trimmed;
        text = applyHardLineBreak(text, li);
        text = ctx.st.pendingQuoteBrCarry + text;

        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        for (int i = 0; i < sp.lines.size(); i++) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(ctx.st.pendingQuoteBrCol);
            setBrSplitLineCell(ctx, cell, sp, i, ctx.styles.normalStyle);
            ctx.st.afterWriteBlockQuoteLine(row.getRowNum(), ctx.st.pendingQuoteBrCol);
        }

        ctx.st.pendingQuoteBr = sp.endsWithBr;
        ctx.st.pendingQuoteBrCarry = sp.carryPrefix;
        return true;
    }

    private static void handleTableSeparatorLine(LineInfo li, RenderContext ctx) {
        ctx.st.afterSkipTableSeparatorLine();
    }

    private static void handleTableRow(LineInfo li, RenderContext ctx) {

        int tableStartCol;
        if (ctx.st.currentTableHeaderRow < 0) {
            tableStartCol = calcBlockStartCol(li.indent, ctx.st);
            ctx.st.currentTableStartCol = tableStartCol;
        } else {
            tableStartCol = ctx.st.currentTableStartCol;
        }

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.TABLE_ROW,
                ctx.styles.normalStyle);

        boolean isHeader = (ctx.st.currentTableHeaderRow < 0);
        int lastCol = MarkdownTable.createTableRow(ctx.wb, li.raw, row, ctx.styles, isHeader, tableStartCol);

        int rowNum = row.getRowNum();
        if (isHeader) {
            ctx.st.currentTableHeaderRow = rowNum;
            ctx.st.currentTableEndCol = lastCol;
            ctx.st.currentTableBodyStartRow = -1;
            ctx.st.currentTableLastBodyRow = -1;
        } else {
            if (ctx.st.currentTableBodyStartRow < 0)
                ctx.st.currentTableBodyStartRow = rowNum;
            ctx.st.currentTableLastBodyRow = rowNum;
            if (lastCol > ctx.st.currentTableEndCol)
                ctx.st.currentTableEndCol = lastCol;
        }

        ctx.st.afterWriteTableRow(tableStartCol);
    }

    private static void handleHeading(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankBeforeHeadingIfNeeded(ctx.sheet, ctx.styles.normalStyle);

        CellStyle style = (li.headingLevel == 1) ? ctx.styles.heading1Style
                : (li.headingLevel == 2) ? ctx.styles.heading2Style
                        : (li.headingLevel == 3) ? ctx.styles.heading3Style : ctx.styles.heading4Style;

        String headingText = applyHardLineBreak(li.headingText, li);
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(headingText);

        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
        Cell cell = row.createCell(0);
        setBrSplitLineCell(ctx, cell, sp, 0, style);
        ctx.st.afterWriteHeading();

        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(0);
            setBrSplitLineCell(ctx, c2, sp, i, style);
            ctx.st.afterWriteHeading();
        }

        ctx.st.pendingHeadingBr = sp.endsWithBr;
        ctx.st.pendingHeadingLevel = li.headingLevel;
        ctx.st.pendingHeadingCarry = sp.carryPrefix;
    }

    private static boolean tryConsumeHeadingBr(LineInfo li, RenderContext ctx) {
        if (!ctx.st.pendingHeadingBr)
            return false;
        if (li.kind != LineKind.NORMAL)
            return false;

        CellStyle style = (ctx.st.pendingHeadingLevel == 1) ? ctx.styles.heading1Style
                : (ctx.st.pendingHeadingLevel == 2) ? ctx.styles.heading2Style
                        : (ctx.st.pendingHeadingLevel == 3) ? ctx.styles.heading3Style : ctx.styles.heading4Style;

        String text = ctx.st.pendingHeadingCarry + applyHardLineBreak(li.trimmed, li);
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        for (int i = 0; i < sp.lines.size(); i++) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(0);
            setBrSplitLineCell(ctx, cell, sp, i, style);
            ctx.st.afterWriteHeading();
        }

        ctx.st.pendingHeadingBr = sp.endsWithBr;
        ctx.st.pendingHeadingCarry = sp.carryPrefix;
        return true;
    }

    private static void handleBullet(LineInfo li, RenderContext ctx) {
        int depth = ListStackUtil.updateListDepth(ctx.st.listStack, li.indent, false);
        int col = clampCol(1 + depth, ctx.st);

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.BULLET_ITEM,
                ctx.styles.normalStyle);

        String bulletText = applyHardLineBreak(li.bulletMarkdownText, li);
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(bulletText);

        Cell cell = row.createCell(col);
        setBrSplitLineCell(ctx, cell, sp, 0, ctx.styles.bulletStyle);
        ctx.st.afterWriteBulletItem(row.getRowNum(), col);

        int contCol = clampCol(col + 1, ctx.st);
        int lastRowNum = row.getRowNum();
        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(contCol);
            setBrSplitLineCell(ctx, c2, sp, i, ctx.styles.bulletStyle);
            lastRowNum = r2.getRowNum();
            ctx.st.afterWriteNormalText(lastRowNum, contCol, 0, false);
        }

        boolean needCont = sp.endsWithBr || sp.lines.size() >= 2;
        if (needCont) {
            ctx.st.pendingListBr = true;
            ctx.st.pendingListBrCol = contCol;
            ctx.st.pendingListBrRow = (sp.lines.size() >= 2) ? lastRowNum : -1;
            ctx.st.pendingListBrHasCell = (sp.lines.size() >= 2);
            ctx.st.pendingListBrStyle = ctx.styles.bulletStyle;
            ctx.st.pendingListBrCarry = sp.carryPrefix;

            ctx.st.bulletDetailActive = false;
        }
    }

    private static void handleNumberedList(LineInfo li, RenderContext ctx) {
        int depth = ListStackUtil.updateListDepth(ctx.st.listStack, li.indent, true);
        int col = clampCol(1 + depth, ctx.st);

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.NUMBER_ITEM,
                ctx.styles.normalStyle);

        String numberedText = applyHardLineBreak(li.trimmed, li);
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(numberedText);

        Cell cell = row.createCell(col);
        setBrSplitLineCell(ctx, cell, sp, 0, ctx.styles.listStyle);
        ctx.st.afterWriteNumberedItem(li.indent, col);

        int contCol = clampCol(col + 1, ctx.st);
        int lastRowNum = row.getRowNum();
        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(contCol);
            setBrSplitLineCell(ctx, c2, sp, i, ctx.styles.listStyle);
            lastRowNum = r2.getRowNum();
            ctx.st.afterWriteNormalText(lastRowNum, contCol, 0, false);
        }

        boolean needCont = sp.endsWithBr || sp.lines.size() >= 2;
        if (needCont) {
            ctx.st.pendingListBr = true;
            ctx.st.pendingListBrCol = contCol;
            ctx.st.pendingListBrRow = (sp.lines.size() >= 2) ? lastRowNum : -1;
            ctx.st.pendingListBrHasCell = (sp.lines.size() >= 2);
            ctx.st.pendingListBrStyle = ctx.styles.listStyle;
            ctx.st.pendingListBrCarry = sp.carryPrefix;

            ctx.st.inNestedNumberBlock = false;
        }
    }

    private static boolean tryConsumeListBr(LineInfo li, RenderContext ctx) {
        if (!ctx.st.pendingListBr)
            return false;

        if (li.kind != LineKind.NORMAL) {
            ctx.st.pendingListBr = false;
            ctx.st.pendingListBrHasCell = false;
            ctx.st.pendingListBrRow = -1;
            ctx.st.pendingListBrCarry = "";
            return false;
        }

        String text = ctx.st.pendingListBrCarry + applyHardLineBreak(li.trimmed, li);
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
        Cell cell = row.createCell(ctx.st.pendingListBrCol);
        setBrSplitLineCell(ctx, cell, sp, 0, ctx.st.pendingListBrStyle);

        ctx.st.pendingListBrRow = row.getRowNum();
        ctx.st.pendingListBrHasCell = true;

        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(ctx.st.pendingListBrCol);
            setBrSplitLineCell(ctx, c2, sp, i, ctx.st.pendingListBrStyle);
            ctx.st.pendingListBrRow = r2.getRowNum();
        }

        ctx.st.pendingListBrCarry = sp.carryPrefix;
        return true;
    }

    private static void handleNormalText(LineInfo li, RenderContext ctx) {

        if (ctx.st.pendingSameColBr) {
            consumePendingSameColBr(li, ctx);
            return;
        }

        String text = applyHardLineBreak(li.trimmed, li);

        if (!MarkdownInline.hasBrOutsideInlineCode(text) && tryAppendToOpenBlockQuote(text, ctx)) {
            return;
        }

        ctx.st.ensureAutoBlankIfPrevHeading(ctx.sheet, ctx.styles.normalStyle);

        int indent = li.indent;

        if (tryAppendToPrevCells(ctx, text, indent)) {
            return;
        }

        NormalTextFlags f = buildNormalTextFlags(indent, ctx.st);
        int col = calcNormalTextCol(indent, ctx.st, f);

        boolean reuseBlank = shouldReuseBlankForNormalText(indent, ctx.st, f);
        Row row = reuseBlank ? RowUtil.reuseLastMarkdownBlankRow(ctx.sheet, ctx.st, ctx.styles.normalStyle)
                : RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        Cell cell = row.createCell(col);
        setBrSplitLineCell(ctx, cell, sp, 0, ctx.styles.normalStyle);
        ctx.st.afterWriteNormalText(row.getRowNum(), col, indent, f.isListNote);

        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(col);
            setBrSplitLineCell(ctx, c2, sp, i, ctx.styles.normalStyle);
            ctx.st.afterWriteNormalText(r2.getRowNum(), col, indent, false);
        }

        if (sp.endsWithBr) {
            ctx.st.pendingSameColBr = true;
            ctx.st.pendingSameColBrCol = col;
            ctx.st.pendingSameColBrStyle = ctx.styles.normalStyle;
            ctx.st.pendingSameColBrCarry = sp.carryPrefix;
        }
    }

    private static void consumePendingSameColBr(LineInfo li, RenderContext ctx) {
        int col = ctx.st.pendingSameColBrCol;
        CellStyle style = ctx.st.pendingSameColBrStyle;

        String text = ctx.st.pendingSameColBrCarry + applyHardLineBreak(li.trimmed, li);
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        for (int i = 0; i < sp.lines.size(); i++) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(col);
            setBrSplitLineCell(ctx, cell, sp, i, style);
            ctx.st.afterWriteNormalText(row.getRowNum(), col, 0, false);
        }

        ctx.st.pendingSameColBr = sp.endsWithBr;
        ctx.st.pendingSameColBrCarry = sp.carryPrefix;

        if (!ctx.st.pendingSameColBr) {
            ctx.st.pendingSameColBrCol = 0;
            ctx.st.pendingSameColBrStyle = null;
            ctx.st.pendingSameColBrCarry = "";
        }
    }

    private static boolean tryConsumeSameColBr(LineInfo li, RenderContext ctx) {
        if (!ctx.st.pendingSameColBr)
            return false;
        if (li.kind != LineKind.NORMAL) {
            ctx.st.pendingSameColBr = false;
            ctx.st.pendingSameColBrCarry = "";
            return false;
        }

        String text = ctx.st.pendingSameColBrCarry + applyHardLineBreak(li.trimmed, li);
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        for (int i = 0; i < sp.lines.size(); i++) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(ctx.st.pendingSameColBrCol);
            setBrSplitLineCell(ctx, cell, sp, i, ctx.st.pendingSameColBrStyle);
            ctx.st.afterWriteNormalText(row.getRowNum(), ctx.st.pendingSameColBrCol, 0, false);
        }

        ctx.st.pendingSameColBr = sp.endsWithBr;
        ctx.st.pendingSameColBrCarry = sp.carryPrefix;
        return true;
    }

    private static boolean tryAppendToOpenBlockQuote(String trimmed, RenderContext ctx) {
        if (!ctx.st.inBlockQuote || ctx.st.blockQuoteCellRow < 0 || ctx.st.blockQuoteCellCol < 0)
            return false;
        if (ctx.st.lastRowType != RenderState.RowType.OTHER || ctx.st.lastBlankFromMarkdown)
            return false;

        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(trimmed);
        appendBrSplitLineWithSpace(ctx, ctx.st.blockQuoteCellRow, ctx.st.blockQuoteCellCol, sp, 0,
                ctx.styles.normalStyle);

        ctx.st.afterAppendToOpenBlockQuoteFromNormalText(ctx.st.blockQuoteCellRow, ctx.st.blockQuoteCellCol);
        return true;
    }

    private static boolean tryAppendToPrevCells(RenderContext ctx, String trimmed, int indent) {

        if (ctx.st.bulletDetailActive && indent > 0 && ctx.st.bulletDetailRow == ctx.st.rowIndex - 1) {
            appendToExistingCellWithBr(ctx, ctx.st.bulletDetailRow, ctx.st.bulletDetailCol, trimmed,
                    ctx.styles.bulletStyle, indent);
            return true;
        }

        boolean isNumberDetail = ctx.st.rowIndex > 0 && ctx.st.inNestedNumberBlock
                && ctx.st.lastContentType == RenderState.ContentType.NUMBER && indent > ctx.st.nestedNumberIndent
                && ctx.st.lastRowType == RenderState.RowType.OTHER && !ctx.st.lastBlankFromMarkdown;

        if (isNumberDetail) {
            int rowNum = ctx.st.rowIndex - 1;
            appendToExistingCellWithBr(ctx, rowNum, ctx.st.nestedNumberCol, trimmed, ctx.styles.listStyle, indent);
            return true;
        }

        boolean isSameIndentConcat = ctx.st.lastContentType == RenderState.ContentType.NORMAL
                && ctx.st.lastNormalRowIndex >= 0 && ctx.st.lastNormalIndent == indent
                && ctx.st.lastRowType == RenderState.RowType.OTHER && !ctx.st.lastBlankFromMarkdown
                && ctx.st.lastNormalRowIndex == ctx.st.rowIndex - 1;

        if (isSameIndentConcat) {
            int rowNum = ctx.st.lastNormalRowIndex;
            int colNum = ctx.st.lastContentCol;
            appendToExistingCellWithBr(ctx, rowNum, colNum, trimmed, ctx.styles.normalStyle, indent);
            return true;
        }

        return false;
    }

    private static void appendToExistingCellWithBr(RenderContext ctx, int targetRow, int targetCol, String markdown,
            CellStyle style, int indent) {

        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(markdown);

        appendBrSplitLineWithSpace(ctx, targetRow, targetCol, sp, 0, style);
        if (!sp.lines.isEmpty()) {
            ctx.st.afterAppendNormalToExistingCell(targetRow, targetCol, indent);
        }

        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(targetCol);
            setBrSplitLineCell(ctx, c2, sp, i, style);
            ctx.st.afterWriteNormalText(r2.getRowNum(), targetCol, indent, false);
        }

        if (sp.endsWithBr) {
            ctx.st.pendingSameColBr = true;
            ctx.st.pendingSameColBrCol = targetCol;
            ctx.st.pendingSameColBrStyle = style;
            ctx.st.pendingSameColBrCarry = sp.carryPrefix;

            ctx.st.resetOnBlockBoundary();
        }
    }

    private static NormalTextFlags buildNormalTextFlags(int indent, RenderState st) {
        boolean isHeadingParagraph = st.inHeadingParagraphBlock && indent == 0 && !st.inListBlock;

        boolean isListNote = st.inListBlock && indent == 0 && st.lastRowType == RenderState.RowType.BLANK
                && st.lastBlankFromMarkdown && st.lastBlankRowIndex >= 0;

        boolean isListChildParagraph = indent > 0 && st.inListBlock && st.lastRowType == RenderState.RowType.BLANK
                && st.lastBlankFromMarkdown && st.lastBlankRowIndex >= 0;

        return new NormalTextFlags(isHeadingParagraph, isListNote, isListChildParagraph);
    }

    private static boolean shouldReuseBlankForNormalText(int indent, RenderState st, NormalTextFlags f) {
        if (st.lastBlankAfterTable) {
            return false;
        }
        boolean collapseEmptyLineBetweenPlainParagraphs = indent == 0 && st.lastRowType == RenderState.RowType.BLANK
                && st.lastBlankFromMarkdown && st.lastBlankRowIndex >= 0 && !st.inListBlock
                && (st.lastContentType == RenderState.ContentType.NORMAL
                        || st.lastContentType == RenderState.ContentType.CODE);

        return f.isListNote || f.isListChildParagraph || collapseEmptyLineBetweenPlainParagraphs;
    }

    private static int calcNormalTextCol(int indent, RenderState st, NormalTextFlags f) {
        if (f.isHeadingParagraph || f.isListNote) {
            return 0;
        }
        if (f.isListChildParagraph) {
            int parentDepth = ListStackUtil.getParentListDepthForChildParagraph(st.listStack);
            int col = 2 + Math.max(0, parentDepth);
            return clampCol(col, st);
        }

        int baseCol;
        if (indent == 0) {
            baseCol = 0;
        } else if (!st.listStack.isEmpty()) {
            int depth = ListStackUtil.getDepthForIndent(st.listStack, indent);
            baseCol = 1 + depth;
        } else {
            int level = indent / 2;
            if (level < 0)
                level = 0;
            baseCol = 1 + level;
        }
        return clampCol(baseCol, st);
    }

    private static int clampCol(int col, RenderState st) {
        if (col < 0)
            return 0;
        if (col >= st.mergeLastCol)
            return st.mergeLastCol - 1;
        return col;
    }

    private static String applyHardLineBreak(String text, LineInfo li) {
        boolean byBackslash = MdTextUtil.hasHardLineBreakByBackslash(li.raw);
        boolean bySpaces = MdTextUtil.hasHardLineBreakBySpaces(li.raw);
        if (!byBackslash && !bySpaces) {
            return text;
        }
        String out = text;
        if (byBackslash) {
            out = MdTextUtil.removeTrailingBackslash(out);
        }
        return out + "<br>";
    }

    private static int calcBlockStartCol(int indent, RenderState st) {
        if (indent <= 0) {
            return 0;
        }

        int col;
        if (!st.listStack.isEmpty()) {
            int depth = ListStackUtil.getDepthForIndent(st.listStack, indent);
            col = 1 + depth;
        } else {
            int level = indent / 2;
            if (level < 0)
                level = 0;
            col = 1 + level;
        }
        return clampCol(col, st);
    }

    private static void setBrSplitLineCell(RenderContext ctx, Cell cell, MarkdownInline.BrSplitResult sp, int lineIndex,
            CellStyle style) {

        List<MarkdownInline.MdSegment> line = (sp.lines.isEmpty() || lineIndex < 0 || lineIndex >= sp.lines.size())
                ? Collections.<MarkdownInline.MdSegment>emptyList()
                : sp.lines.get(lineIndex);

        MarkdownInline.setResolvedSegmentsCell(ctx.wb, cell, line, style);
    }

    private static void appendBrSplitLineWithSpace(RenderContext ctx, int rowNum, int colNum,
            MarkdownInline.BrSplitResult sp, int lineIndex, CellStyle style) {

        if (sp.lines.isEmpty() || lineIndex < 0 || lineIndex >= sp.lines.size()) {
            return;
        }

        CellAppendUtil.appendResolvedSegmentsWithSpace(ctx, rowNum, colNum, sp.lines.get(lineIndex), style);
    }

    private static final class NormalTextFlags {
        final boolean isHeadingParagraph;
        final boolean isListNote;
        final boolean isListChildParagraph;

        NormalTextFlags(boolean hp, boolean ln, boolean lcp) {
            this.isHeadingParagraph = hp;
            this.isListNote = ln;
            this.isListChildParagraph = lcp;
        }
    }
}
