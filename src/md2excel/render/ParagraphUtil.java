package md2excel.render;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import md2excel.markdown.ListStackUtil;
import md2excel.markdown.MdTextUtil;

final class ParagraphUtil {

    private ParagraphUtil() {
    }

    static boolean isParagraphLine(MarkdownRenderer.LineInfo li) {
        if (li == null) {
            return false;
        }

        if (li.kind == MarkdownRenderer.LineKind.BLOCK_QUOTE) {
            MarkdownRenderer.LineInfo q = li.quotedContent;
            if (q == null) {
                return false;
            }

            return q.kind == MarkdownRenderer.LineKind.BULLET_ITEM || q.kind == MarkdownRenderer.LineKind.NUMBER_ITEM
                    || isQuotedNormalParagraphKind(q.kind);
        }

        switch (li.kind) {
        case NORMAL:
        case BULLET_ITEM:
        case NUMBER_ITEM:
            return true;

        default:
            return false;
        }
    }

    private static boolean isQuotedNormalParagraphKind(MarkdownRenderer.LineKind kind) {

        switch (kind) {
        case BLANK:
        case HEADING:
        case BULLET_ITEM:
        case NUMBER_ITEM:
        case TABLE_SEPARATOR:
        case TABLE_ROW:
            return false;

        default:
            return true;
        }
    }

    static boolean canContinue(ParagraphBuffer p, MarkdownRenderer.LineInfo li) {

        if (p == null || li == null) {
            return false;
        }

        switch (p.kind) {
        case NORMAL:
            return li.kind == MarkdownRenderer.LineKind.NORMAL;

        case QUOTE_NORMAL:
            return isQuotedNormalParagraphLine(li);

        case BULLET:
            return li.kind == MarkdownRenderer.LineKind.NORMAL && li.indent > p.baseIndent;

        case NUMBER:
            return li.kind == MarkdownRenderer.LineKind.NORMAL && li.indent > p.baseIndent;

        case QUOTE_BULLET:
        case QUOTE_NUMBER:
            return isQuotedNormalParagraphLine(li) && li.quotedContent.indent > p.baseIndent;

        default:
            return false;
        }
    }

    private static boolean isQuotedNormalParagraphLine(MarkdownRenderer.LineInfo li) {

        return li.kind == MarkdownRenderer.LineKind.BLOCK_QUOTE && li.quotedContent != null
                && isQuotedNormalParagraphKind(li.quotedContent.kind);
    }

    static ParagraphBuffer start(MarkdownRenderer.LineInfo li, RenderContext ctx) {

        if (li.kind == MarkdownRenderer.LineKind.BLOCK_QUOTE) {
            MarkdownRenderer.LineInfo q = li.quotedContent;

            if (q == null) {
                throw new IllegalArgumentException("Quoted content is missing");
            }

            switch (q.kind) {
            case BULLET_ITEM:
                return startQuoteBullet(li, ctx);

            case NUMBER_ITEM:
                return startQuoteNumber(li, ctx);

            default:
                if (isQuotedNormalParagraphKind(q.kind)) {
                    return startQuoteNormal(li, ctx);
                }

                throw new IllegalArgumentException("Unsupported quote paragraph kind: " + q.kind);
            }
        }

        switch (li.kind) {
        case NORMAL:
            return startNormal(li, ctx);

        case BULLET_ITEM:
            return startBullet(li, ctx);

        case NUMBER_ITEM:
            return startNumber(li, ctx);

        default:
            throw new IllegalArgumentException("Unsupported paragraph line kind: " + li.kind);
        }
    }

    static void append(ParagraphBuffer p, MarkdownRenderer.LineInfo li) {
        if (p == null || li == null) {
            return;
        }

        String lineText = extractContinuationLineText(p, li);
        p.appendLine(normalizeInlineLineText(lineText), li.endsWithHardBreak);
    }

    static void flush(ParagraphBuffer p, RenderContext ctx) {
        if (p == null || p.isEmpty()) {
            return;
        }

        List<List<MarkdownInline.MdSegment>> lines = parseParagraphToDisplayLines(p.getParagraphText());
        lines = prependPrefixToFirstLine(lines, p.firstLinePrefix);
        lines = ensureAtLeastOneDisplayLine(lines);

        Row firstRow = createFirstRow(p, ctx);
        writeLine(ctx, firstRow, p.firstCol, p.firstLineStyle, lines.get(0));
        afterWriteFirstLine(p, ctx, firstRow.getRowNum());

        for (int i = 1; i < lines.size(); i++) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            writeLine(ctx, row, p.continuationCol, p.continuationStyle, lines.get(i));
            afterWriteContinuationLine(p, ctx, row.getRowNum(), p.continuationCol);
        }
    }

    // ------------------------------------------------------------
    // start
    // ------------------------------------------------------------

    private static ParagraphBuffer startNormal(MarkdownRenderer.LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankIfPrevHeading(ctx.sheet, ctx.styles.blankRowStyle);

        NormalTextFlags f = buildNormalTextFlags(li.indent, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.NORMAL);
        p.baseIndent = li.indent;
        p.firstCol = calcNormalTextCol(li.indent, ctx.st, f);
        p.continuationCol = p.firstCol;
        p.firstLineStyle = ctx.styles.normalStyle;
        p.continuationStyle = ctx.styles.normalStyle;
        p.reuseMarkdownBlankForFirstRow = shouldReuseBlankForNormalText(ctx.st, f);
        p.isListNote = f.isListNote;

        p.appendLine(normalizeInlineLineText(li.paragraphText), li.endsWithHardBreak);
        return p;
    }

    private static ParagraphBuffer startBullet(MarkdownRenderer.LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankBeforeChildListIfNeeded(ctx.sheet, ctx.styles.blankRowStyle, li.indent);

        int depth = ListStackUtil.updateListDepth(ctx.st.listStack, li.indent, false);
        int col = clampCol(ctx.st.startColIndex + 1 + depth, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.BULLET);
        p.baseIndent = li.indent;
        p.firstCol = col;
        p.continuationCol = clampCol(col + 1, ctx.st);
        p.firstLineStyle = ctx.styles.bulletStyle;
        p.continuationStyle = ctx.styles.bulletStyle;
        p.firstLinePrefix = (li.listMarkerText == null) ? "・ " : li.listMarkerText;

        p.appendLine(normalizeInlineLineText(li.listContentText), li.endsWithHardBreak);
        return p;
    }

    private static ParagraphBuffer startNumber(MarkdownRenderer.LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankBeforeChildListIfNeeded(ctx.sheet, ctx.styles.blankRowStyle, li.indent);

        int depth = ListStackUtil.updateListDepth(ctx.st.listStack, li.indent, true);
        int col = clampCol(ctx.st.startColIndex + 1 + depth, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.NUMBER);
        p.baseIndent = li.indent;
        p.firstCol = col;
        p.continuationCol = clampCol(col + 1, ctx.st);
        p.firstLineStyle = ctx.styles.listStyle;
        p.continuationStyle = ctx.styles.listStyle;
        p.firstLinePrefix = (li.listMarkerText == null) ? "" : li.listMarkerText;

        p.appendLine(normalizeInlineLineText(li.listContentText), li.endsWithHardBreak);
        return p;
    }

    private static ParagraphBuffer startQuoteNormal(MarkdownRenderer.LineInfo li, RenderContext ctx) {

        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        MarkdownRenderer.LineInfo q = li.quotedContent;

        int quoteStartCol = calcQuoteStartCol(li.indent, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.QUOTE_NORMAL);

        p.baseIndent = q.indent;
        p.inBlockQuote = true;
        p.quoteStartCol = quoteStartCol;
        p.quoteDecorCol = clampCol(quoteStartCol - 1, ctx.st);
        p.firstCol = quoteStartCol;
        p.continuationCol = quoteStartCol;
        p.firstLineStyle = ctx.styles.normalStyle;
        p.continuationStyle = ctx.styles.normalStyle;

        p.appendLine(normalizeInlineLineText(quotedNormalText(q)), q.endsWithHardBreak);

        return p;
    }

    private static ParagraphBuffer startQuoteBullet(MarkdownRenderer.LineInfo li, RenderContext ctx) {

        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        MarkdownRenderer.LineInfo q = li.quotedContent;

        int quoteStartCol = calcQuoteStartCol(li.indent, ctx.st);

        ensureQuotedAutoBlankBeforeChildListIfNeeded(li, ctx, quoteStartCol);

        int depth = ListStackUtil.updateListDepth(ctx.st.listStack, q.indent, false);

        int col = clampCol(quoteStartCol + 1 + depth, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.QUOTE_BULLET);

        p.baseIndent = q.indent;
        p.inBlockQuote = true;
        p.quoteStartCol = quoteStartCol;
        p.quoteDecorCol = clampCol(quoteStartCol - 1, ctx.st);
        p.firstCol = col;
        p.continuationCol = clampCol(col + 1, ctx.st);
        p.firstLineStyle = ctx.styles.bulletStyle;
        p.continuationStyle = ctx.styles.bulletStyle;
        p.firstLinePrefix = (q.listMarkerText == null) ? "・ " : q.listMarkerText;

        p.appendLine(normalizeInlineLineText(q.listContentText), q.endsWithHardBreak);

        return p;
    }

    private static ParagraphBuffer startQuoteNumber(MarkdownRenderer.LineInfo li, RenderContext ctx) {

        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        MarkdownRenderer.LineInfo q = li.quotedContent;

        int quoteStartCol = calcQuoteStartCol(li.indent, ctx.st);

        ensureQuotedAutoBlankBeforeChildListIfNeeded(li, ctx, quoteStartCol);

        int depth = ListStackUtil.updateListDepth(ctx.st.listStack, q.indent, true);

        int col = clampCol(quoteStartCol + 1 + depth, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.QUOTE_NUMBER);

        p.baseIndent = q.indent;
        p.inBlockQuote = true;
        p.quoteStartCol = quoteStartCol;
        p.quoteDecorCol = clampCol(quoteStartCol - 1, ctx.st);
        p.firstCol = col;
        p.continuationCol = clampCol(col + 1, ctx.st);
        p.firstLineStyle = ctx.styles.listStyle;
        p.continuationStyle = ctx.styles.listStyle;
        p.firstLinePrefix = (q.listMarkerText == null) ? "" : q.listMarkerText;

        p.appendLine(normalizeInlineLineText(q.listContentText), q.endsWithHardBreak);

        return p;
    }

    // ------------------------------------------------------------
    // flush helpers
    // ------------------------------------------------------------

    private static Row createFirstRow(ParagraphBuffer p, RenderContext ctx) {
        if (p.reuseMarkdownBlankForFirstRow) {
            return RowUtil.reuseLastMarkdownBlankRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
        }
        return RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
    }

    private static void writeLine(RenderContext ctx, Row row, int col, CellStyle style,
            List<MarkdownInline.MdSegment> segments) {

        Cell cell = row.createCell(col);
        MarkdownInline.setResolvedSegmentsCell(ctx.wb, cell, segments, style);
    }

    private static void afterWriteFirstLine(ParagraphBuffer p, RenderContext ctx, int rowNum) {
        switch (p.kind) {
        case NORMAL:
            ctx.st.afterWriteNormalText(rowNum, p.firstCol, p.baseIndent, p.isListNote);
            break;

        case BULLET:
            ctx.st.afterWriteBulletItem(rowNum, p.firstCol);
            break;

        case NUMBER:
            ctx.st.afterWriteNumberedItem(p.baseIndent, p.firstCol);
            break;

        case QUOTE_NORMAL:
            ctx.st.afterWriteNormalText(rowNum, p.firstCol, p.baseIndent, false);
            recordQuotedRow(ctx, rowNum, p.quoteStartCol, p.firstCol);
            break;

        case QUOTE_BULLET:
            ctx.st.afterWriteBulletItem(rowNum, p.firstCol);
            recordQuotedRow(ctx, rowNum, p.quoteStartCol, -1);
            break;

        case QUOTE_NUMBER:
            ctx.st.afterWriteNumberedItem(p.baseIndent, p.firstCol);
            recordQuotedRow(ctx, rowNum, p.quoteStartCol, -1);
            break;

        default:
            break;
        }
    }

    private static void afterWriteContinuationLine(ParagraphBuffer p, RenderContext ctx, int rowNum, int col) {
        switch (p.kind) {
        case NORMAL:
        case BULLET:
        case NUMBER:
            ctx.st.afterWriteNormalText(rowNum, col, p.baseIndent, false);
            break;

        case QUOTE_NORMAL:
            ctx.st.afterWriteNormalText(rowNum, col, p.baseIndent, false);
            recordQuotedRow(ctx, rowNum, p.quoteStartCol, col);
            break;

        case QUOTE_BULLET:
        case QUOTE_NUMBER:
            ctx.st.afterWriteNormalText(rowNum, col, p.baseIndent, false);
            recordQuotedRow(ctx, rowNum, p.quoteStartCol, -1);
            break;

        default:
            break;
        }
    }

    private static List<List<MarkdownInline.MdSegment>> prependPrefixToFirstLine(
            List<List<MarkdownInline.MdSegment>> lines, String prefix) {

        if (prefix == null || prefix.isEmpty()) {
            return (lines == null) ? Collections.<List<MarkdownInline.MdSegment>>emptyList() : lines;
        }

        List<List<MarkdownInline.MdSegment>> out = new ArrayList<List<MarkdownInline.MdSegment>>();

        if (lines == null || lines.isEmpty()) {
            List<MarkdownInline.MdSegment> first = new ArrayList<MarkdownInline.MdSegment>();
            first.add(new MarkdownInline.MdSegment(prefix, false, false, false));
            out.add(first);
            return out;
        }

        List<MarkdownInline.MdSegment> first = new ArrayList<MarkdownInline.MdSegment>();
        first.add(new MarkdownInline.MdSegment(prefix, false, false, false));
        first.addAll(lines.get(0));
        out.add(first);

        for (int i = 1; i < lines.size(); i++) {
            out.add(lines.get(i));
        }

        return out;
    }

    private static List<List<MarkdownInline.MdSegment>> ensureAtLeastOneDisplayLine(
            List<List<MarkdownInline.MdSegment>> lines) {

        if (lines != null && !lines.isEmpty()) {
            return lines;
        }

        List<List<MarkdownInline.MdSegment>> out = new ArrayList<List<MarkdownInline.MdSegment>>();
        out.add(Collections.<MarkdownInline.MdSegment>emptyList());
        return out;
    }

    /**
     * 想定する MarkdownInline 側の追加 API:
     *
     * static List<List<MdSegment>> parseParagraphToDisplayLines(String
     * paragraphText)
     *
     * - paragraphText 全体を1回だけ inline 解析する - ParagraphBuffer.SOFT_BREAK_TOKEN は空白扱い
     * - ParagraphBuffer.HARD_BREAK_TOKEN で表示行を分割 - 戻り値は「表示行ごとの resolved segments」
     */
    private static List<List<MarkdownInline.MdSegment>> parseParagraphToDisplayLines(String paragraphText) {
        return MarkdownInline.parseParagraphToDisplayLines(paragraphText);
    }

    // ------------------------------------------------------------
    // text extract / normalize
    // ------------------------------------------------------------

    private static String extractContinuationLineText(ParagraphBuffer p, MarkdownRenderer.LineInfo li) {

        switch (p.kind) {
        case NORMAL:
        case BULLET:
        case NUMBER:
            return li.paragraphText;

        case QUOTE_NORMAL:
        case QUOTE_BULLET:
        case QUOTE_NUMBER:
            return quotedNormalText(li.quotedContent);

        default:
            return "";
        }
    }

    private static String normalizeInlineLineText(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }

        return MdTextUtil.replaceBrOutsideInlineCode(text, String.valueOf(ParagraphBuffer.HARD_BREAK_TOKEN));
    }

    // ------------------------------------------------------------
    // quoted blank helper
    // ------------------------------------------------------------

    private static void ensureQuotedAutoBlankBeforeChildListIfNeeded(MarkdownRenderer.LineInfo li, RenderContext ctx,
            int quoteStartCol) {

        if (li.quotedContent != null && ctx.st.shouldInsertAutoBlankBeforeChildList(li.quotedContent.indent)) {

            writeQuotedBlankRow(ctx, quoteStartCol);
        }
    }

    private static void writeQuotedBlankRow(RenderContext ctx, int quoteStartCol) {
        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.blankRowStyle);

        Cell cell = row.createCell(quoteStartCol);
        MarkdownInline.setResolvedSegmentsCell(ctx.wb, cell, Collections.<MarkdownInline.MdSegment>emptyList(),
                ctx.styles.blankRowStyle);

        ctx.st.blankBlockQuoteRows.add(row.getRowNum());

        recordQuotedRow(ctx, row.getRowNum(), quoteStartCol, -1);

        ctx.st.lastRowType = RenderState.RowType.BLANK;
        ctx.st.lastLineWasTable = false;
        ctx.st.lastBlankFromMarkdown = false;
        ctx.st.lastBlankRowIndex = -1;
        ctx.st.lastBlankAfterTable = false;

        ctx.st.lastContentType = RenderState.ContentType.NORMAL;
        ctx.st.lastContentCol = quoteStartCol;
        ctx.st.lastContentWasTable = false;

        ctx.st.lastNormalRowIndex = -1;
        ctx.st.lastNormalIndent = -1;
        ctx.st.bulletDetailActive = false;
        ctx.st.lastWasBlockQuote = true;
    }

    private static void recordQuotedRow(RenderContext ctx, int rowNum, int quoteStartCol, int appendCellCol) {
        int quoteDecorCol = clampCol(quoteStartCol - 1, ctx.st);

        if (!ctx.st.inBlockQuote || ctx.st.blockQuoteFirstRow < 0) {
            ctx.st.inBlockQuote = true;
            ctx.st.blockQuoteFirstRow = rowNum;
            ctx.st.blockQuoteCol = quoteDecorCol;
        }

        if (quoteDecorCol < ctx.st.blockQuoteCol) {
            ctx.st.blockQuoteCol = quoteDecorCol;
        }

        ctx.st.blockQuoteLastRow = rowNum;

        if (appendCellCol >= 0) {
            ctx.st.blockQuoteCellRow = rowNum;
            ctx.st.blockQuoteCellCol = appendCellCol;
        } else {
            ctx.st.blockQuoteCellRow = -1;
            ctx.st.blockQuoteCellCol = -1;
        }

        ctx.st.lastWasBlockQuote = true;
    }

    // ------------------------------------------------------------
    // normal text placement
    // ------------------------------------------------------------

    private static final class NormalTextFlags {
        final boolean isHeadingParagraph;
        final boolean isListNote;
        final boolean isListChildParagraph;

        NormalTextFlags(boolean headingParagraph, boolean listNote, boolean listChildParagraph) {
            this.isHeadingParagraph = headingParagraph;
            this.isListNote = listNote;
            this.isListChildParagraph = listChildParagraph;
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

    private static boolean shouldReuseBlankForNormalText(RenderState st, NormalTextFlags f) {
        if (st.lastBlankAfterTable) {
            return false;
        }
        return f.isListChildParagraph;
    }

    private static int calcNormalTextCol(int indent, RenderState st, NormalTextFlags f) {
        if (f.isHeadingParagraph || f.isListNote) {
            return rootCol(st);
        }

        if (f.isListChildParagraph) {
            int parentDepth = ListStackUtil.getParentListDepthForChildParagraph(st.listStack);
            int col = st.startColIndex + 2 + Math.max(0, parentDepth);
            return clampCol(col, st);
        }

        int baseCol;
        if (indent == 0) {
            baseCol = st.startColIndex;
        } else if (!st.listStack.isEmpty()) {
            int depth = ListStackUtil.getDepthForIndent(st.listStack, indent);
            baseCol = st.startColIndex + 1 + depth;
        } else {
            int level = indent / 2;
            if (level < 0) {
                level = 0;
            }
            baseCol = st.startColIndex + 1 + level;
        }

        return clampCol(baseCol, st);
    }

    private static String quotedNormalText(MarkdownRenderer.LineInfo q) {

        if (q == null) {
            return "";
        }

        if (q.kind == MarkdownRenderer.LineKind.NORMAL) {
            return q.paragraphText;
        }

        String text = q.trimmed;

        if (MdTextUtil.hasHardLineBreakByBackslash(q.raw)) {
            text = MdTextUtil.removeTrailingBackslash(text);
        }

        return text;
    }

    // ------------------------------------------------------------
    // col helpers
    // ------------------------------------------------------------

    private static int calcQuoteStartCol(int indent, RenderState st) {
        return clampCol(calcBlockStartCol(indent, st) + 1, st);
    }

    private static int calcBlockStartCol(int indent, RenderState st) {
        if (indent <= 0) {
            return rootCol(st);
        }

        int col;
        if (!st.listStack.isEmpty()) {
            int depth = ListStackUtil.getDepthForIndent(st.listStack, indent);
            col = st.startColIndex + 1 + depth;
        } else {
            int level = indent / 2;
            if (level < 0) {
                level = 0;
            }
            col = st.startColIndex + 1 + level;
        }

        return clampCol(col, st);
    }

    private static int rootCol(RenderState st) {
        return clampCol(st.startColIndex, st);
    }

    private static int clampCol(int col, RenderState st) {
        if (col < 0) {
            return 0;
        }
        if (col >= st.mergeLastCol) {
            return st.mergeLastCol - 1;
        }
        return col;
    }
}