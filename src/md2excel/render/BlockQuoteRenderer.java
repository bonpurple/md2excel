package md2excel.render;

import static md2excel.render.RenderLayout.calcQuoteStartCol;

import md2excel.markdown.LineInfo;
import md2excel.markdown.LineKind;

final class BlockQuoteRenderer {

    private BlockQuoteRenderer() {
    }

    static void render(LineInfo line, RenderContext ctx) {

        if (!line.isQuoted()) {
            throw new IllegalArgumentException("Quoted line was expected");
        }

        int quoteStartCol = calcQuoteStartCol(line.getIndent(), ctx.st);

        int quoteDepth = line.getQuoteDepth();

        /*
         * 引用内コードブロック中は、通常の引用境界処理を行わない。
         */
        if (ctx.st.codeBlock().isInBlockQuote()) {
            renderInsideQuotedCodeBlock(line, quoteStartCol, quoteDepth, ctx);

            return;
        }

        boolean explicitBlankAfterQuotedCode = line.getKind() == LineKind.BLANK
                && ctx.st.isLastContentType(RenderState.ContentType.CODE) && ctx.st.wasLastBlockQuote();

        if (!explicitBlankAfterQuotedCode) {
            ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        }

        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        switch (line.getKind()) {
        case BLANK:
            renderBlank(quoteStartCol, quoteDepth, ctx);
            return;

        case HORIZONTAL_RULE:
            HorizontalRuleRenderer.renderQuoted(ctx, quoteStartCol, quoteDepth);
            return;

        case HEADING:
            HeadingRenderer.renderQuoted(line, quoteStartCol, quoteDepth, ctx);
            return;

        case CODE_FENCE:
            CodeBlockRenderer.renderQuotedFence(line, quoteStartCol, quoteDepth, ctx);
            return;

        case TABLE_SEPARATOR:
            TableRenderer.renderQuotedSeparator(ctx);
            return;

        case TABLE_ROW:
            TableRenderer.renderQuotedRow(line, quoteStartCol, quoteDepth, ctx);
            return;

        case NORMAL:
        case BULLET_ITEM:
        case NUMBER_ITEM:
            throw new AssertionError(
                    "Quote paragraph line should have been " + "handled by ParagraphUtil: " + line.getKind());

        case CODE_LINE:
            throw new AssertionError("CODE_LINE outside quoted code block");

        default:
            throw new AssertionError("Unhandled quoted line: " + line.getKind());
        }
    }

    private static void renderInsideQuotedCodeBlock(LineInfo line, int quoteStartCol, int quoteDepth,
            RenderContext ctx) {

        switch (line.getKind()) {
        case CODE_LINE:
            CodeBlockRenderer.renderQuotedLine(line, ctx);
            return;

        case CODE_FENCE:
            CodeBlockRenderer.renderQuotedFence(line, quoteStartCol, quoteDepth, ctx);
            return;

        default:
            throw new AssertionError("Unexpected line inside quoted code block: " + line.getKind());
        }
    }

    private static void renderBlank(int quoteStartCol, int quoteDepth, RenderContext ctx) {

        ctx.st.resetOnBlockBoundary();
        ctx.st.leaveListBlockPreservingLevels();

        if (ctx.st.isLastRowType(RenderState.RowType.HORIZONTAL_RULE) && ctx.st.wasLastBlockQuote()) {

            ctx.st.afterConsumeQuotedMarkdownBlankWithoutNewRow();
            return;
        }

        BlockQuoteRowUtil.writeBlankRow(ctx, quoteStartCol, quoteDepth);
    }
}