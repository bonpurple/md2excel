package md2excel.render;

import static md2excel.render.RenderLayout.clampCol;
import static md2excel.render.RenderLayout.rootCol;

import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import md2excel.markdown.LineInfo;

final class HeadingRenderer {

    private HeadingRenderer() {
    }

    static void render(LineInfo line, RenderContext ctx) {

        ctx.st.ensureAutoBlankBeforeHeadingIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        CellStyle style = resolveStyle(line.getHeadingLevel(), ctx);

        List<List<MarkdownInline.MdSegment>> lines = parseDisplayLines(line.getHeadingText());

        for (int index = 0; index < lines.size(); index++) {

            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

            Cell cell = row.createCell(rootCol(ctx.st));

            MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, lines.get(index), style);

            ctx.st.afterWriteHeading();
        }
    }

    static void renderQuoted(LineInfo line, int quoteStartCol, int quoteDepth, RenderContext ctx) {

        CellStyle style = resolveStyle(line.getHeadingLevel(), ctx);

        List<List<MarkdownInline.MdSegment>> lines = parseDisplayLines(line.getHeadingText());

        int textCol = clampCol(quoteStartCol + quoteDepth - 1, ctx.st);

        RenderState.QuoteRowKind rowKind = toQuoteRowKind(line.getHeadingLevel());

        for (int index = 0; index < lines.size(); index++) {

            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

            Cell cell = row.createCell(textCol);

            MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, lines.get(index), style);

            ctx.st.afterWriteQuotedHeading();

            ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, textCol, rowKind, quoteDepth);
        }
    }

    private static List<List<MarkdownInline.MdSegment>> parseDisplayLines(String text) {

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(text);

        if (!lines.isEmpty()) {
            return lines;
        }

        return Collections.<List<MarkdownInline.MdSegment>>singletonList(
                Collections.<MarkdownInline.MdSegment>emptyList());
    }

    private static CellStyle resolveStyle(int headingLevel, RenderContext ctx) {

        switch (headingLevel) {
        case 1:
            return ctx.styles.heading1Style;

        case 2:
            return ctx.styles.heading2Style;

        case 3:
            return ctx.styles.heading3Style;

        default:
            return ctx.styles.heading4Style;
        }
    }

    private static RenderState.QuoteRowKind toQuoteRowKind(int headingLevel) {

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