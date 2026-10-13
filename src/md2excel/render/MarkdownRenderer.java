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

            boolean tableParsingEnabled = TableParsingProbe.shouldEnableTableParsing(rawLine, cursor.peek(), state);

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

            MdBlockBoundary.processLineBoundary(line, ctx);

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

        MdBlockBoundary.finishAtEof(paragraph, ctx);
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
}
