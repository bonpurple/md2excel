package md2excel.render;

import static md2excel.render.RenderLayout.calcQuoteStartCol;
import static md2excel.render.RenderLayout.clampCol;
import static md2excel.render.RenderLayout.rootCol;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import md2excel.markdown.LineInfo;
import md2excel.markdown.LineKind;
import md2excel.markdown.MdTextUtil;

final class ParagraphUtil {

    private ParagraphUtil() {
    }

    static boolean isParagraphLine(LineInfo line) {
        if (line == null) {
            return false;
        }

        if (line.isQuoted()) {
            return line.getKind() == LineKind.BULLET_ITEM || line.getKind() == LineKind.NUMBER_ITEM
                    || isQuotedNormalParagraphKind(line.getKind());
        }

        switch (line.getKind()) {
        case NORMAL:
        case BULLET_ITEM:
        case NUMBER_ITEM:
            return true;

        default:
            return false;
        }
    }

    private static boolean isQuotedNormalParagraphKind(LineKind kind) {

        return kind == LineKind.NORMAL;
    }

    static boolean canContinue(ParagraphBuffer p, LineInfo li) {

        if (p == null || li == null) {
            return false;
        }

        switch (p.kind) {
        case NORMAL:
            return !li.isQuoted() && li.getKind() == LineKind.NORMAL;

        case QUOTE_NORMAL:
            return isQuotedNormalParagraphLine(li) && li.getQuoteDepth() == p.quoteDepth;

        case BULLET:
        case NUMBER:
            return !li.isQuoted() && li.getKind() == LineKind.NORMAL && li.getContentIndent() > p.baseIndent;

        case QUOTE_BULLET:
        case QUOTE_NUMBER:
            return isQuotedNormalParagraphLine(li) && li.getQuoteDepth() == p.quoteDepth
                    && li.getContentIndent() > p.baseIndent;

        default:
            return false;
        }
    }

    private static boolean isQuotedNormalParagraphLine(LineInfo line) {

        return line != null && line.isQuoted() && isQuotedNormalParagraphKind(line.getKind());
    }

    static ParagraphBuffer start(LineInfo li, RenderContext ctx) {

        if (li.isQuoted()) {
            switch (li.getKind()) {
            case BULLET_ITEM:
                return startQuoteBullet(li, ctx);

            case NUMBER_ITEM:
                return startQuoteNumber(li, ctx);

            case NORMAL:
                return startQuoteNormal(li, ctx);

            default:
                throw new IllegalArgumentException("Unsupported quote paragraph kind: " + li.getKind());
            }
        }

        switch (li.getKind()) {
        case NORMAL:
            return startNormal(li, ctx);

        case BULLET_ITEM:
            return startBullet(li, ctx);

        case NUMBER_ITEM:
            return startNumber(li, ctx);

        default:
            throw new IllegalArgumentException("Unsupported paragraph line kind: " + li.getKind());
        }
    }

    static void append(ParagraphBuffer p, LineInfo li) {
        if (p == null || li == null) {
            return;
        }

        String lineText = extractContinuationLineText(p, li);
        p.appendLine(lineText, li.endsWithHardBreak());
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

    private static ParagraphBuffer startNormal(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankIfPrevHeading(ctx.sheet, ctx.styles.blankRowStyle);

        NormalTextFlags f = buildNormalTextFlags(li.getIndent(), ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.NORMAL);
        p.baseIndent = li.getIndent();
        p.firstCol = calcNormalTextCol(li.getIndent(), ctx.st, f);
        p.continuationCol = p.firstCol;
        p.firstLineStyle = ctx.styles.normalStyle;
        p.continuationStyle = ctx.styles.normalStyle;
        p.reuseMarkdownBlankForFirstRow = shouldReuseBlankForNormalText(ctx.st, f);
        p.isListNote = f.isListNote;

        p.appendLine(li.getParagraphText(), li.endsWithHardBreak());
        return p;
    }

    private static ParagraphBuffer startBullet(LineInfo li, RenderContext ctx) {
        return startNormalList(li, ctx, false);
    }

    private static ParagraphBuffer startNumber(LineInfo li, RenderContext ctx) {
        return startNormalList(li, ctx, true);
    }

    private static ParagraphBuffer startNormalList(LineInfo li, RenderContext ctx, boolean ordered) {
        ctx.st.ensureAutoBlankBeforeChildListIfNeeded(ctx.sheet, ctx.styles.blankRowStyle, li.getIndent());

        int depth = ctx.st.updateListDepth(li.getIndent(), ordered);

        int col = clampCol(ctx.st.getStartColIndex() + 1 + depth, ctx.st);

        ParagraphBuffer.Kind kind = ordered ? ParagraphBuffer.Kind.NUMBER : ParagraphBuffer.Kind.BULLET;
        CellStyle style = ordered ? ctx.styles.listStyle : ctx.styles.bulletStyle;
        String defaultMarker = ordered ? "" : "・ ";

        ParagraphBuffer p = new ParagraphBuffer(kind);
        p.baseIndent = li.getIndent();
        p.firstCol = col;
        p.continuationCol = clampCol(col + 1, ctx.st);
        p.firstLineStyle = style;
        p.continuationStyle = style;
        p.firstLinePrefix = (li.getListMarkerText() == null) ? defaultMarker : li.getListMarkerText();

        p.appendLine(li.getListContentText(), li.endsWithHardBreak());
        return p;
    }

    private static ParagraphBuffer startQuoteNormal(LineInfo li, RenderContext ctx) {

        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);

        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        int quoteStartCol = calcQuoteStartCol(li.getIndent(), ctx.st);

        int quoteDepth = li.getQuoteDepth();

        int textCol = clampCol(quoteStartCol + quoteDepth - 1, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.QUOTE_NORMAL);

        p.baseIndent = li.getContentIndent();
        p.quoteStartCol = quoteStartCol;
        p.quoteDepth = quoteDepth;

        p.firstCol = textCol;
        p.continuationCol = textCol;

        p.firstLineStyle = ctx.styles.normalStyle;
        p.continuationStyle = ctx.styles.normalStyle;

        p.appendLine(quotedNormalText(li), li.endsWithHardBreak());

        return p;
    }

    private static ParagraphBuffer startQuoteBullet(LineInfo li, RenderContext ctx) {

        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        int quoteStartCol = calcQuoteStartCol(li.getIndent(), ctx.st);
        int quoteDepth = Math.max(1, li.getQuoteDepth());

        ensureQuotedAutoBlankBeforeChildListIfNeeded(li, ctx, quoteStartCol, quoteDepth);

        int listDepth = ctx.st.updateListDepth(li.getContentIndent(), false);

        int col = clampCol(quoteStartCol + quoteDepth + listDepth, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.QUOTE_BULLET);

        p.baseIndent = li.getContentIndent();
        p.quoteStartCol = quoteStartCol;
        p.quoteDepth = quoteDepth;

        p.firstCol = col;
        p.continuationCol = clampCol(col + 1, ctx.st);

        p.firstLineStyle = ctx.styles.bulletStyle;
        p.continuationStyle = ctx.styles.bulletStyle;
        p.firstLinePrefix = li.getListMarkerText() == null ? "・ " : li.getListMarkerText();

        p.appendLine(li.getListContentText(), li.endsWithHardBreak());

        return p;
    }

    private static ParagraphBuffer startQuoteNumber(LineInfo li, RenderContext ctx) {

        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        int quoteStartCol = calcQuoteStartCol(li.getIndent(), ctx.st);
        int quoteDepth = Math.max(1, li.getQuoteDepth());

        ensureQuotedAutoBlankBeforeChildListIfNeeded(li, ctx, quoteStartCol, quoteDepth);

        int listDepth = ctx.st.updateListDepth(li.getContentIndent(), true);

        int col = clampCol(quoteStartCol + quoteDepth + listDepth, ctx.st);

        ParagraphBuffer p = new ParagraphBuffer(ParagraphBuffer.Kind.QUOTE_NUMBER);

        p.baseIndent = li.getContentIndent();
        p.quoteStartCol = quoteStartCol;
        p.quoteDepth = quoteDepth;

        p.firstCol = col;
        p.continuationCol = clampCol(col + 1, ctx.st);

        p.firstLineStyle = ctx.styles.listStyle;
        p.continuationStyle = ctx.styles.listStyle;
        p.firstLinePrefix = li.getListMarkerText() == null ? "" : li.getListMarkerText();

        p.appendLine(li.getListContentText(), li.endsWithHardBreak());

        return p;
    }

    static int getSetextHeadingLevel(ParagraphBuffer p, LineInfo li) {

        if (p == null || li == null || p.isEmpty()) {
            return 0;
        }

        // 今回は通常 paragraph の Setext 化だけを対象にする。
        // 引用・リスト内はそれぞれの対応時に拡張する。
        if (p.kind != ParagraphBuffer.Kind.NORMAL) {
            return 0;
        }

        // CommonMark: heading content の先頭行は最大3スペース。
        if (p.baseIndent > 3) {
            return 0;
        }

        return parseSetextUnderlineLevel(li.getRaw());
    }

    private static int parseSetextUnderlineLevel(String raw) {
        if (raw == null || raw.isEmpty()) {
            return 0;
        }

        int n = raw.length();
        int i = 0;
        int leadingSpaces = 0;

        while (i < n && raw.charAt(i) == ' ') {
            leadingSpaces++;
            if (leadingSpaces > 3) {
                return 0;
            }
            i++;
        }

        // marker 前の tab は4-column境界になるので、
        // この簡易判定では Setext underline としない。
        if (i < n && raw.charAt(i) == '\t') {
            return 0;
        }

        if (i >= n) {
            return 0;
        }

        char marker = raw.charAt(i);
        if (marker != '=' && marker != '-') {
            return 0;
        }

        while (i < n && raw.charAt(i) == marker) {
            i++;
        }

        // marker の後ろは space / tab のみ許可。
        while (i < n) {
            char ch = raw.charAt(i);
            if (ch != ' ' && ch != '\t') {
                return 0;
            }
            i++;
        }

        return marker == '=' ? 1 : 2;
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
        MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, segments, style);
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

            ctx.st.recordBlockQuoteRow(rowNum, p.quoteStartCol, p.firstCol, RenderState.QuoteRowKind.NORMAL,
                    p.quoteDepth);

            break;

        case QUOTE_BULLET:
            ctx.st.afterWriteBulletItem(rowNum, p.firstCol);

            ctx.st.recordBlockQuoteRow(rowNum, p.quoteStartCol, p.firstCol, RenderState.QuoteRowKind.NORMAL,
                    p.quoteDepth);

            break;

        case QUOTE_NUMBER:
            ctx.st.afterWriteNumberedItem(p.baseIndent, p.firstCol);

            ctx.st.recordBlockQuoteRow(rowNum, p.quoteStartCol, p.firstCol, RenderState.QuoteRowKind.NORMAL,
                    p.quoteDepth);

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

            ctx.st.recordBlockQuoteRow(rowNum, p.quoteStartCol, col, RenderState.QuoteRowKind.NORMAL, p.quoteDepth);

            break;

        case QUOTE_BULLET:
        case QUOTE_NUMBER:
            ctx.st.afterWriteNormalText(rowNum, col, p.baseIndent, false);

            ctx.st.recordBlockQuoteRow(rowNum, p.quoteStartCol, col, RenderState.QuoteRowKind.NORMAL, p.quoteDepth);

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

    static void flushSetextHeading(ParagraphBuffer p, int headingLevel, RenderContext ctx) {

        if (p == null || p.isEmpty()) {
            return;
        }

        List<List<MarkdownInline.MdSegment>> lines = parseParagraphToDisplayLines(p.getParagraphText());

        lines = prependPrefixToFirstLine(lines, p.firstLinePrefix);
        lines = ensureAtLeastOneDisplayLine(lines);

        CellStyle headingStyle = headingLevel == 1 ? ctx.styles.heading1Style : ctx.styles.heading2Style;

        for (int i = 0; i < lines.size(); i++) {
            Row row = (i == 0) ? createFirstRow(p, ctx) : RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

            writeLine(ctx, row, p.firstCol, headingStyle, lines.get(i));
        }

        // 通常の見出しと同じ状態にする。
        // 次の通常段落開始時に自動空行が1行入る。
        ctx.st.afterWriteHeading();
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
    // text extract
    // ------------------------------------------------------------

    private static String extractContinuationLineText(ParagraphBuffer p, LineInfo li) {

        switch (p.kind) {
        case NORMAL:
        case BULLET:
        case NUMBER:
            return li.getParagraphText();

        case QUOTE_NORMAL:
        case QUOTE_BULLET:
        case QUOTE_NUMBER:
            return quotedNormalText(li);

        default:
            return "";
        }
    }

    // ------------------------------------------------------------
    // quoted blank helper
    // ------------------------------------------------------------

    private static void ensureQuotedAutoBlankBeforeChildListIfNeeded(LineInfo li, RenderContext ctx, int quoteStartCol,
            int quoteDepth) {

        if (ctx.st.shouldInsertAutoBlankBeforeChildList(li.getContentIndent())) {

            BlockQuoteRowUtil.writeBlankRow(ctx, quoteStartCol, quoteDepth);
        }
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

        boolean isHeadingParagraph = st.isInHeadingParagraphBlock() && indent == 0 && !st.isInListBlock();

        boolean hasReusableBlank = st.hasReusableMarkdownBlankForParagraph();

        boolean isListNote = st.isInListBlock() && indent == 0 && hasReusableBlank;

        boolean isListChildParagraph = indent > 0 && st.isInListBlock() && hasReusableBlank;

        return new NormalTextFlags(isHeadingParagraph, isListNote, isListChildParagraph);
    }

    private static boolean shouldReuseBlankForNormalText(RenderState st, NormalTextFlags f) {

        if (st.isLastBlankAfterTable()) {
            return false;
        }

        return f.isListChildParagraph;
    }

    private static int calcNormalTextCol(int indent, RenderState st, NormalTextFlags f) {
        if (f.isHeadingParagraph || f.isListNote) {
            return rootCol(st);
        }

        if (f.isListChildParagraph) {
            int parentDepth = st.getParentListDepthForChildParagraph();

            int col = st.getStartColIndex() + 2 + Math.max(0, parentDepth);

            return clampCol(col, st);
        }

        int baseCol;

        if (indent == 0) {
            baseCol = st.getStartColIndex();

        } else if (st.hasListLevels()) {
            int depth = st.getListDepthForIndent(indent);

            baseCol = st.getStartColIndex() + 1 + depth;

        } else {
            int level = Math.max(0, indent / 2);

            baseCol = st.getStartColIndex() + 1 + level;
        }

        return clampCol(baseCol, st);
    }

    private static String quotedNormalText(LineInfo line) {

        if (line == null) {
            return "";
        }

        if (line.getKind() == LineKind.NORMAL) {
            return line.getParagraphText();
        }

        String text = line.getContentTrimmed();

        if (MdTextUtil.hasHardLineBreakByBackslash(line.getContentRaw())) {

            text = MdTextUtil.removeTrailingBackslash(text);
        }

        return text;
    }
}
