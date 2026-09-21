package md2excel.render;

import static md2excel.render.RenderLayout.calcBlockStartCol;
import static md2excel.render.RenderLayout.calcQuoteStartCol;
import static md2excel.render.RenderLayout.clampCol;
import static md2excel.render.RenderLayout.rootCol;

import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import md2excel.excel.ExcelCellUtil;
import md2excel.excel.Md2ExcelSheetUtil;
import md2excel.markdown.MdTextUtil;

public final class MarkdownRenderer {

    static enum LineKind {
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

    static final class LineInfo {
        final String raw;
        final String trimmed;
        final int indent;
        final LineKind kind;

        final int headingLevel;
        final String headingText;

        final boolean endsWithHardBreak;
        final String paragraphText;
        final String listMarkerText;
        final String listContentText;

        // kind == BLOCK_QUOTE のときだけ設定。
        // 引用マーカーを除去した内容を通常と同じ parseContent() で解析した結果。
        final LineInfo quotedContent;

        LineInfo(String raw, String trimmed, int indent, LineKind kind, int headingLevel, String headingText,
                boolean endsWithHardBreak, String paragraphText, String listMarkerText, String listContentText,
                LineInfo quotedContent) {

            this.raw = raw;
            this.trimmed = trimmed;
            this.indent = indent;
            this.kind = kind;

            this.headingLevel = headingLevel;
            this.headingText = headingText;

            this.endsWithHardBreak = endsWithHardBreak;
            this.paragraphText = paragraphText;
            this.listMarkerText = listMarkerText;
            this.listContentText = listContentText;

            this.quotedContent = quotedContent;
        }

        boolean isTableLike() {
            if (kind == LineKind.TABLE_SEPARATOR || kind == LineKind.TABLE_ROW) {
                return true;
            }

            if (kind != LineKind.BLOCK_QUOTE) {
                return false;
            }

            LineInfo content = getInnermostQuotedContent();

            return content != null && (content.kind == LineKind.TABLE_SEPARATOR || content.kind == LineKind.TABLE_ROW);
        }

        int getQuoteDepth() {
            int depth = 0;
            LineInfo current = this;

            while (current != null && current.kind == LineKind.BLOCK_QUOTE) {
                depth++;
                current = current.quotedContent;
            }

            return depth;
        }

        LineInfo getInnermostQuotedContent() {
            LineInfo current = this;

            while (current != null && current.kind == LineKind.BLOCK_QUOTE) {
                current = current.quotedContent;
            }

            return current;
        }
    }

    public static void render(Iterator<String> it, RenderContext ctx) {
        RenderState st = ctx.st;
        ParagraphBuffer para = null;

        LineCursor cursor = new LineCursor(it);

        while (cursor.hasNext()) {
            String rawLine = cursor.next();

            boolean tableParsingEnabled = shouldEnableTableParsing(rawLine, cursor.peek(), st);

            LineInfo li = MarkdownLineParser.parse(rawLine, st, tableParsingEnabled);

            // Setext heading は直前の paragraph と現在行をセットで判定する。
            // underline 行自体は Excel 行として出力しない。
            int setextHeadingLevel = ParagraphUtil.getSetextHeadingLevel(para, li);

            if (setextHeadingLevel > 0) {
                ParagraphUtil.flushSetextHeading(para, setextHeadingLevel, ctx);
                para = null;
                continue;
            }

            // まず open paragraph が継続できるか判定
            if (para != null && ParagraphUtil.canContinue(para, li)) {
                ParagraphUtil.append(para, li);
                continue;
            }

            // 継続できないなら flush
            if (para != null) {
                ParagraphUtil.flush(para, ctx);
                para = null;
            }

            MdBlockBoundary.closeTableIfLeaving(li.isTableLike(), ctx);
            MdBlockBoundary.apply(li.kind.policy, ctx);

            // paragraph 対象行は start して次へ
            if (ParagraphUtil.isParagraphLine(li)) {
                para = ParagraphUtil.start(li, ctx);
                continue;
            }

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
                handleBlockQuote(li, ctx); // quote blank のみ到達
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
            case NUMBER_ITEM:
            case NORMAL:
                throw new AssertionError("Paragraph line should have been handled earlier: " + li.kind);
            default:
                throw new AssertionError("Unhandled LineKind: " + li.kind);
            }
        }

        if (para != null) {
            ParagraphUtil.flush(para, ctx);
        }

        if (st.codeBlock().isOpen()) {
            finishCodeBlock(ctx);
        }

        if (st.lastLineWasTable) {
            MarkdownTable.closeTableIfOpen(ctx.sheet, ctx.styles, st);
        }
        BlockQuoteUtil.closeBlockQuoteIfOpen(ctx.sheet, ctx.styles, st);
    }

    private static void handleCodeFence(LineInfo li, RenderContext ctx) {

        CodeBlockState codeBlock = ctx.st.codeBlock();

        // 開始
        if (!codeBlock.isOpen()) {
            ctx.st.ensureAutoBlankIfPrevBlockQuote(ctx.sheet, ctx.styles.blankRowStyle);

            codeBlock.open(MdTextUtil.getCodeFenceMarker(li.trimmed), MdTextUtil.getCodeFenceLength(li.trimmed),
                    li.indent, false, -1);

            ctx.st.lastLineWasTable = false;
            return;
        }

        // 終了
        finishCodeBlock(ctx);
    }

    private static void finishCodeBlock(RenderContext ctx) {

        CodeBlockState codeBlock = ctx.st.codeBlock();

        if (codeBlock.hasRenderedLines()) {
            int fillEndCol = Math.max(codeBlock.getStartCol(), ctx.st.renderLastColIndex);

            for (int r = codeBlock.getFirstRow(); r <= codeBlock.getLastRow(); r++) {

                Row rowObj = ctx.sheet.getRow(r);
                if (rowObj == null) {
                    continue;
                }

                for (int c = codeBlock.getStartCol(); c <= fillEndCol; c++) {

                    Cell cell = ExcelCellUtil.getOrCreateCell(rowObj, c);

                    boolean isTop = r == codeBlock.getFirstRow();
                    boolean isBottom = r == codeBlock.getLastRow();
                    boolean isLeft = c == codeBlock.getStartCol();
                    boolean isRight = c == fillEndCol;

                    int mask = 0;

                    if (isTop) {
                        mask |= 1;
                    }
                    if (isBottom) {
                        mask |= 2;
                    }
                    if (isLeft) {
                        mask |= 4;
                    }
                    if (isRight) {
                        mask |= 8;
                    }

                    cell.setCellStyle(ctx.styles.codeBlockFrameStyle(mask));
                }
            }
        }

        codeBlock.reset();
        ctx.st.lastLineWasTable = false;
    }

    private static void handleInCodeBlock(LineInfo li, RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.CODE_LINE,
                ctx.styles.normalStyle);

        int openingIndent = ctx.st.codeBlock().getOpeningIndent();

        // 装飾はブロック開始列から、コード本文は1列右へ配置
        int frameStartCol = calcBlockStartCol(openingIndent, ctx.st);
        int codeCol = clampCol(frameStartCol + 1, ctx.st);

        String codeLine = MdTextUtil.removeLeadingIndentColumns(li.raw, openingIndent);
        codeLine = MdTextUtil.expandTabs(codeLine);

        Cell cell = row.createCell(codeCol);
        MarkdownInline.setCodeBlockRichTextCell(ctx.fontCache, cell, codeLine, ctx.styles.codeBlockStyle);

        ctx.st.codeBlock().recordLine(row.getRowNum(), frameStartCol);

        ctx.st.afterWriteCodeLine(codeCol);
    }

    private static void handleBlankLine(LineInfo li, RenderContext ctx) {
        ctx.st.onMarkdownBlankLine(ctx.sheet, ctx.styles.blankRowStyle);
    }

    private static void handleHorizontalRule(LineInfo li, RenderContext ctx) {
        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx, RowUtil.ReuseKind.HORIZONTAL_RULE,
                ctx.styles.blankRowStyle);
        Md2ExcelSheetUtil.createHorizontalRuleRow(ctx.sheet, row, ctx.styles.horizontalRuleStyle, ctx.st.startColIndex,
                ctx.st.renderEndColExclusive);
        ctx.st.afterWriteHorizontalRule();
    }

    private static void handleBlockQuote(LineInfo li, RenderContext ctx) {

        int quoteStartCol = calcQuoteStartCol(li.indent, ctx.st);
        int quoteDepth = Math.max(1, li.getQuoteDepth());

        LineInfo q = li.getInnermostQuotedContent();

        if (q == null) {
            throw new AssertionError("Quoted content is missing");
        }

        // すでに引用内コードブロック中なら、
        // 通常の block quote 前後処理を通さない。
        if (ctx.st.codeBlock().isInBlockQuote()) {

            switch (q.kind) {
            case CODE_LINE:
                handleQuotedCodeLine(q, quoteStartCol, ctx);
                return;

            case CODE_FENCE:
                handleQuotedCodeFence(q, quoteStartCol, ctx);
                return;

            default:
                throw new AssertionError("Unexpected line inside quoted code block: " + q.kind);
            }
        }

        // quoted code の直後に明示的な `>` 空行がある場合、
        // code block 用の自動空行は追加しない。
        boolean explicitBlankAfterQuotedCode = q.kind == LineKind.BLANK
                && ctx.st.lastContentType == RenderState.ContentType.CODE && ctx.st.lastWasBlockQuote;

        if (!explicitBlankAfterQuotedCode) {
            ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        }

        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        switch (q.kind) {
        case BLANK:
            handleQuotedBlank(ctx, quoteStartCol, quoteDepth);
            break;

        case HORIZONTAL_RULE:
            handleQuotedHorizontalRule(ctx, quoteStartCol, quoteDepth);
            break;

        case HEADING:
            handleQuotedHeading(q, quoteStartCol, quoteDepth, ctx);
            break;

        case CODE_FENCE:
            handleQuotedCodeFence(q, quoteStartCol, ctx);
            break;

        case TABLE_SEPARATOR:
            ctx.st.afterSkipTableSeparatorLine();
            ctx.st.lastWasBlockQuote = true;
            break;

        case TABLE_ROW:
            int tableStartCol = clampCol(quoteStartCol + quoteDepth - 1, ctx.st);

            MarkdownTable.TableRowRenderResult rr = renderTableRow(q.raw, tableStartCol, quoteDepth, ctx);

            for (int r = rr.firstRowNum; r <= rr.lastRowNum; r++) {
                ctx.st.recordBlockQuoteTableRow(r, quoteStartCol, tableStartCol, rr.lastCol, quoteDepth,
                        rr.getStyleRole(r));
            }

            ctx.st.lastWasBlockQuote = true;
            break;

        default:
            throw new AssertionError("Quote paragraph line should have been handled by ParagraphUtil: " + q.kind);
        }
    }

    private static void handleTableSeparatorLine(LineInfo li, RenderContext ctx) {
        ctx.st.afterSkipTableSeparatorLine();
    }

    private static void handleTableRow(LineInfo li, RenderContext ctx) {
        renderTableRow(li.raw, calcBlockStartCol(li.indent, ctx.st), 0, ctx);
    }

    private static MarkdownTable.TableRowRenderResult renderTableRow(String tableLine, int firstRowStartCol,
            int quoteDepth, RenderContext ctx) {

        TableState table = ctx.st.table();

        boolean isHeader = !table.isOpen();

        int tableStartCol = isHeader ? firstRowStartCol : table.getStartCol();

        MarkdownTable.TableRowRenderResult result = MarkdownTable.createTableRows(ctx, tableLine, isHeader,
                tableStartCol);

        if (isHeader) {
            table.begin(result.firstRowNum, tableStartCol, result.lastCol, quoteDepth);
        } else {
            table.recordBodyRows(result.firstRowNum, result.lastRowNum, result.lastCol);
        }

        ctx.st.afterWriteTableRow(tableStartCol);
        return result;
    }

    private static void handleHeading(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankBeforeHeadingIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        CellStyle style = resolveHeadingStyle(li.headingLevel, ctx);

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(li.headingText);
        if (lines.isEmpty()) {
            lines = Collections.<List<MarkdownInline.MdSegment>>singletonList(
                    Collections.<MarkdownInline.MdSegment>emptyList());
        }

        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
        Cell cell = row.createCell(rootCol(ctx.st));
        MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, lines.get(0), style);
        ctx.st.afterWriteHeading();

        for (int i = 1; i < lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(rootCol(ctx.st));
            MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, c2, lines.get(i), style);
            ctx.st.afterWriteHeading();
        }
    }

    private static void handleQuotedBlank(RenderContext ctx, int quoteStartCol, int quoteDepth) {

        ctx.st.resetOnBlockBoundary();
        ctx.st.clearListContext();

        // 通常Markdown空行と同じく、
        // 水平線直後の空行は Excel 行を増やさない。
        if (ctx.st.lastRowType == RenderState.RowType.HORIZONTAL_RULE && ctx.st.lastWasBlockQuote) {

            ctx.st.afterConsumeMarkdownBlankWithoutNewRow();
            ctx.st.lastWasBlockQuote = true;
            return;
        }

        BlockQuoteRowUtil.writeBlankRow(ctx, quoteStartCol, quoteDepth);
    }

    private static void handleQuotedHorizontalRule(RenderContext ctx, int quoteStartCol, int quoteDepth) {

        int previousRowNum = ctx.st.rowIndex - 1;

        boolean reusePreviousQuotedBlank = ctx.st.lastRowType == RenderState.RowType.BLANK && previousRowNum >= 0
                && ctx.st.isBlockQuoteRowKind(previousRowNum, RenderState.QuoteRowKind.BLANK);

        Row row;

        if (reusePreviousQuotedBlank) {
            row = ctx.sheet.getRow(previousRowNum);

            if (row == null) {
                row = ctx.sheet.createRow(previousRowNum);
            }
        } else {
            row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.blankRowStyle);

            int contentCol = clampCol(quoteStartCol + quoteDepth - 1, ctx.st);

            Cell cell = row.createCell(contentCol);

            MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell,
                    Collections.<MarkdownInline.MdSegment>emptyList(), ctx.styles.blankRowStyle);
        }

        // 再利用した行も含め、行種別と引用深度を上書きする。
        ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, -1, RenderState.QuoteRowKind.HORIZONTAL_RULE,
                quoteDepth);

        ctx.st.afterWriteHorizontalRule();

        // afterWriteHorizontalRule()は通常水平線として
        // lastWasBlockQuote=falseにするため、引用状態へ戻す。
        ctx.st.lastWasBlockQuote = true;
    }

    private static void handleQuotedHeading(LineInfo q, int quoteStartCol, int quoteDepth, RenderContext ctx) {

        CellStyle style = resolveHeadingStyle(q.headingLevel, ctx);

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(q.headingText);

        if (lines.isEmpty()) {
            lines = Collections.<List<MarkdownInline.MdSegment>>singletonList(
                    Collections.<MarkdownInline.MdSegment>emptyList());
        }

        int textCol = clampCol(quoteStartCol + quoteDepth - 1, ctx.st);
        RenderState.QuoteRowKind quoteRowKind = toQuoteHeadingRowKind(q.headingLevel);

        for (int i = 0; i < lines.size(); i++) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

            Cell cell = row.createCell(textCol);

            MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, lines.get(i), style);

            ctx.st.afterWriteQuotedHeading(textCol);

            ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, textCol, quoteRowKind, quoteDepth);
        }
    }

    private static void handleQuotedCodeFence(LineInfo q, int quoteStartCol, RenderContext ctx) {

        CodeBlockState codeBlock = ctx.st.codeBlock();

        // 開始
        if (!codeBlock.isOpen()) {
            codeBlock.open(MdTextUtil.getCodeFenceMarker(q.trimmed), MdTextUtil.getCodeFenceLength(q.trimmed), q.indent,
                    true, quoteStartCol);

            ctx.st.lastLineWasTable = false;
            ctx.st.lastWasBlockQuote = true;
            return;
        }

        // 終了
        finishCodeBlock(ctx);

        // finishCodeBlock()は通常コードブロックとして終了するため、
        // 引用コンテキストだけ戻す。
        ctx.st.lastWasBlockQuote = true;
    }

    private static void handleQuotedCodeLine(LineInfo q, int quoteStartCol, RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.CODE_LINE,
                ctx.styles.normalStyle);

        // B列: quote decoration
        // C列: code block frame
        // D列: code text
        int frameStartCol = quoteStartCol;
        int codeCol = clampCol(frameStartCol + 1, ctx.st);

        int trimColumns = ctx.st.codeBlock().getOpeningIndent();

        String codeLine = MdTextUtil.removeLeadingIndentColumns(q.raw, trimColumns);
        codeLine = MdTextUtil.expandTabs(codeLine);

        Cell cell = row.createCell(codeCol);

        MarkdownInline.setCodeBlockRichTextCell(ctx.fontCache, cell, codeLine, ctx.styles.codeBlockStyle);

        ctx.st.codeBlock().recordLine(row.getRowNum(), frameStartCol);

        // 引用終了時にコードstyleを上書きしないため記録。
        ctx.st.afterWriteCodeLine(codeCol);

        ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, -1, RenderState.QuoteRowKind.CODE);

        ctx.st.lastWasBlockQuote = true;
    }

    private static CellStyle resolveHeadingStyle(int headingLevel, RenderContext ctx) {

        return (headingLevel == 1) ? ctx.styles.heading1Style
                : (headingLevel == 2) ? ctx.styles.heading2Style
                        : (headingLevel == 3) ? ctx.styles.heading3Style : ctx.styles.heading4Style;
    }

    private static final class LineCursor {
        private final Iterator<String> source;
        private boolean hasBuffered;
        private String buffered;

        LineCursor(Iterator<String> source) {
            this.source = source;
            advance();
        }

        boolean hasNext() {
            return hasBuffered;
        }

        String next() {
            String current = buffered;
            advance();
            return current;
        }

        String peek() {
            return hasBuffered ? buffered : null;
        }

        private void advance() {
            hasBuffered = source.hasNext();
            buffered = hasBuffered ? source.next() : null;
        }
    }

    private static final class TableProbeLine {
        final int quoteDepth;
        final String content;

        TableProbeLine(int quoteDepth, String content) {
            this.quoteDepth = quoteDepth;
            this.content = content;
        }
    }

    private static TableProbeLine unwrapQuoteMarkers(String rawLine) {
        int quoteDepth = 0;
        String content = rawLine;

        while (content.trim().startsWith(">")) {
            content = MarkdownLineParser.stripOneQuoteMarker(content);
            quoteDepth++;
        }

        return new TableProbeLine(quoteDepth, content);
    }

    private static boolean shouldEnableTableParsing(String rawLine, String nextRawLine, RenderState st) {

        TableProbeLine current = unwrapQuoteMarkers(rawLine);

        boolean continuingTable = st.lastLineWasTable && st.table().getQuoteDepth() == current.quoteDepth;

        if (continuingTable) {
            return true;
        }

        if (nextRawLine == null) {
            return false;
        }

        TableProbeLine next = unwrapQuoteMarkers(nextRawLine);

        return current.quoteDepth == next.quoteDepth && MarkdownTable.isTableStart(current.content, next.content);
    }

    private static RenderState.QuoteRowKind toQuoteHeadingRowKind(int headingLevel) {

        switch (headingLevel) {
        case 1:
            return RenderState.QuoteRowKind.HEADING_1;

        case 2:
            return RenderState.QuoteRowKind.HEADING_2;

        case 3:
            return RenderState.QuoteRowKind.HEADING_3;

        default:
            return RenderState.QuoteRowKind.HEADING_4;
        }
    }
}