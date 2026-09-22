package md2excel.render;

import java.util.Iterator;

import md2excel.markdown.LineInfo;

public final class MarkdownRenderer {

    private MarkdownRenderer() {
    }

    public static void render(Iterator<String> source, RenderContext ctx) {

        RenderState state = ctx.st;
        ParagraphBuffer paragraph = null;

        LineCursor cursor = new LineCursor(source);

        while (cursor.hasNext()) {
            String rawLine = cursor.next();

            boolean tableParsingEnabled = shouldEnableTableParsing(rawLine, cursor.peek(), state);

            LineInfo line = MarkdownLineParser.parse(rawLine, state, tableParsingEnabled);

            /*
             * Setext見出しは、直前の段落と現在行を セットで判定する。
             */
            int setextHeadingLevel = ParagraphUtil.getSetextHeadingLevel(paragraph, line);

            if (setextHeadingLevel > 0) {
                ParagraphUtil.flushSetextHeading(paragraph, setextHeadingLevel, ctx);

                paragraph = null;
                continue;
            }

            if (paragraph != null && ParagraphUtil.canContinue(paragraph, line)) {

                ParagraphUtil.append(paragraph, line);

                continue;
            }

            if (paragraph != null) {
                ParagraphUtil.flush(paragraph, ctx);

                paragraph = null;
            }

            MdBlockBoundary.closeTableIfLeaving(line, ctx);

            MdBlockBoundary.apply(MdBlockBoundary.policyFor(line), ctx);

            if (ParagraphUtil.isParagraphLine(line)) {
                paragraph = ParagraphUtil.start(line, ctx);

                continue;
            }

            if (line.isQuoted()) {
                BlockQuoteRenderer.render(line, ctx);

                continue;
            }

            renderBlockLine(line, ctx);
        }

        /*
         * EOFでは、保留段落、未閉鎖コード枠、表、引用装飾をこの順序で確定する。
         * 引用装飾は引用段落の出力と表の最終行役割の確定後に行い、後続行用の自動空行挿入や通常のブロック境界状態リセットは行わない。
         */
        if (paragraph != null) {
            ParagraphUtil.flush(paragraph, ctx);
        }

        if (state.codeBlock().isOpen()) {
            CodeBlockRenderer.finish(ctx);
        }

        if (state.isLastLineTable()) {
            MarkdownTable.closeTableIfOpen(ctx.sheet, ctx.styles, state);
        }

        BlockQuoteUtil.closeBlockQuoteIfOpen(ctx.sheet, ctx.styles, state);
    }

    private static void renderBlockLine(LineInfo line, RenderContext ctx) {

        switch (line.getKind()) {
        case CODE_FENCE:
            CodeBlockRenderer.renderFence(line, ctx);
            return;

        case CODE_LINE:
            CodeBlockRenderer.renderLine(line, ctx);
            return;

        case BLANK:
            ctx.st.onMarkdownBlankLine(ctx.sheet, ctx.styles.blankRowStyle);
            return;

        case HORIZONTAL_RULE:
            HorizontalRuleRenderer.render(ctx);
            return;

        case TABLE_SEPARATOR:
            TableRenderer.renderSeparator(ctx);
            return;

        case TABLE_ROW:
            TableRenderer.renderRow(line, ctx);
            return;

        case HEADING:
            HeadingRenderer.render(line, ctx);
            return;

        case BULLET_ITEM:
        case NUMBER_ITEM:
        case NORMAL:
            throw new AssertionError("Paragraph line should have been " + "handled earlier: " + line.getKind());

        default:
            throw new AssertionError("Unhandled LineKind: " + line.getKind());
        }
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

    private static boolean shouldEnableTableParsing(String rawLine, String nextRawLine, RenderState state) {

        TableProbeLine current = unwrapQuoteMarkers(rawLine);

        boolean continuingTable = state.isLastLineTable() && state.table().getQuoteDepth() == current.quoteDepth;

        if (continuingTable) {
            return true;
        }

        if (nextRawLine == null) {
            return false;
        }

        TableProbeLine next = unwrapQuoteMarkers(nextRawLine);

        return current.quoteDepth == next.quoteDepth && MarkdownTable.isTableStart(current.content, next.content);
    }
}
