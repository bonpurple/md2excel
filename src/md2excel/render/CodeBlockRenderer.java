package md2excel.render;

import static md2excel.render.RenderLayout.calcBlockStartCol;
import static md2excel.render.RenderLayout.clampCol;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import md2excel.excel.CodeBlockFrameMask;
import md2excel.excel.ExcelCellUtil;
import md2excel.markdown.CodeFence;
import md2excel.markdown.LineInfo;
import md2excel.markdown.MdTextUtil;

final class CodeBlockRenderer {

    private CodeBlockRenderer() {
    }

    static void renderFence(LineInfo line, RenderContext ctx) {

        CodeBlockState codeBlock = ctx.st.codeBlock();

        // 開始フェンス
        if (!codeBlock.isOpen()) {
            ctx.st.ensureAutoBlankIfPrevBlockQuote(ctx.sheet, ctx.styles.blankRowStyle);

            CodeFence fence = requireOpeningFence(line);

            codeBlock.open(fence.getMarker(), fence.getLength(), line.getContentIndent(), false, -1, 0);

            ctx.st.afterOpenNormalCodeFence();
            return;
        }

        // 終了フェンス
        finish(ctx);
    }

    static void renderLine(LineInfo line, RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.CODE_LINE,
                ctx.styles.normalStyle);

        CodeBlockState codeBlock = ctx.st.codeBlock();

        int openingIndent = codeBlock.getOpeningIndent();

        int frameStartCol = calcBlockStartCol(openingIndent, ctx.st);

        int codeCol = clampCol(frameStartCol + 1, ctx.st);

        writeCodeLine(line, row, codeCol, openingIndent, ctx);

        codeBlock.recordLine(row.getRowNum(), frameStartCol);

        ctx.st.afterWriteCodeLine(codeCol);
    }

    static void renderQuotedFence(LineInfo line, int quoteStartCol, int quoteDepth, RenderContext ctx) {

        CodeBlockState codeBlock = ctx.st.codeBlock();

        // 開始フェンス
        if (!codeBlock.isOpen()) {
            CodeFence fence = requireOpeningFence(line);

            codeBlock.open(fence.getMarker(), fence.getLength(), line.getContentIndent(), true, quoteStartCol,
                    quoteDepth);

            ctx.st.afterOpenQuotedCodeFence();
            return;
        }

        // 終了フェンス
        finish(ctx);
    }

    static void renderQuotedLine(LineInfo line, RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.CODE_LINE,
                ctx.styles.normalStyle);

        CodeBlockState codeBlock = ctx.st.codeBlock();

        int quoteStartCol = codeBlock.getQuoteStartCol();

        int quoteDepth = codeBlock.getQuoteDepth();

        /*
         * quoteStartCol - 1からquoteDepth列分を引用装飾列とし、 その右側にコード枠とコード本文を配置する。
         */
        int frameStartCol = RenderLayout.calcQuoteContentCol(quoteStartCol, quoteDepth, ctx.st);

        int codeCol = clampCol(frameStartCol + 1, ctx.st);

        writeCodeLine(line, row, codeCol, codeBlock.getOpeningIndent(), ctx);

        codeBlock.recordLine(row.getRowNum(), frameStartCol);

        ctx.st.afterWriteQuotedCodeLine(codeCol);

        ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, -1, RenderState.QuoteRowKind.CODE, quoteDepth);
    }

    private static void writeCodeLine(LineInfo line, Row row, int codeCol, int openingIndent, RenderContext ctx) {

        String codeLine = MdTextUtil.removeLeadingIndentColumns(line.getContentRaw(), openingIndent);

        codeLine = MdTextUtil.expandTabs(codeLine);

        Cell cell = row.createCell(codeCol);

        MarkdownInline.setCodeBlockRichTextCell(ctx.fontCache, cell, codeLine, ctx.styles.codeBlockStyle);
    }

    static void finish(RenderContext ctx) {

        CodeBlockState codeBlock = ctx.st.codeBlock();

        boolean quoted = codeBlock.isInBlockQuote();

        if (codeBlock.hasRenderedLines()) {
            applyFrame(ctx, codeBlock);
        }

        codeBlock.reset();
        ctx.st.afterFinishCodeBlock(quoted);
    }

    private static void applyFrame(RenderContext ctx, CodeBlockState codeBlock) {

        int fillEndCol = Math.max(codeBlock.getStartCol(), ctx.st.getRenderLastColIndex());

        for (int rowIndex = codeBlock.getFirstRow(); rowIndex <= codeBlock.getLastRow(); rowIndex++) {

            Row row = ctx.sheet.getRow(rowIndex);

            if (row == null) {
                continue;
            }

            for (int col = codeBlock.getStartCol(); col <= fillEndCol; col++) {

                Cell cell = ExcelCellUtil.getOrCreateCell(row, col);

                int mask = createFrameMask(rowIndex, col, fillEndCol, codeBlock);

                cell.setCellStyle(ctx.styles.codeBlockFrameStyle(mask));
            }
        }
    }

    private static int createFrameMask(int rowIndex, int col, int fillEndCol, CodeBlockState codeBlock) {

        int mask = CodeBlockFrameMask.NONE;

        if (rowIndex == codeBlock.getFirstRow()) {
            mask |= CodeBlockFrameMask.TOP;
        }

        if (rowIndex == codeBlock.getLastRow()) {
            mask |= CodeBlockFrameMask.BOTTOM;
        }

        if (col == codeBlock.getStartCol()) {
            mask |= CodeBlockFrameMask.LEFT;
        }

        if (col == fillEndCol) {
            mask |= CodeBlockFrameMask.RIGHT;
        }

        return mask;
    }

    private static CodeFence requireOpeningFence(LineInfo line) {

        CodeFence fence = CodeFence.parseOpening(line.getContentTrimmed());

        if (fence == null) {
            throw new IllegalArgumentException("Opening code fence was expected: " + line.getContentRaw());
        }

        return fence;
    }
}
