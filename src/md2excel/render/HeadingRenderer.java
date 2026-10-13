package md2excel.render;

import static md2excel.render.RenderLayout.calcQuoteContentCol;
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

        List<List<MdSegment>> lines = parseDisplayLines(line.getHeadingText());

        for (int index = 0; index < lines.size(); index++) {

            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

            Cell cell = row.createCell(rootCol(ctx.st));

            MarkdownRichTextWriter.setResolvedSegmentsCell(ctx.fontCache, cell, lines.get(index), style);

            ctx.st.afterWriteHeading();
        }
    }

    static void renderQuoted(LineInfo line, int quoteStartCol, int quoteDepth, RenderContext ctx) {

        CellStyle style = resolveStyle(line.getHeadingLevel(), ctx);

        List<List<MdSegment>> lines = parseDisplayLines(line.getHeadingText());

        int textCol = calcQuoteContentCol(quoteStartCol, quoteDepth, ctx.st);

        QuoteRowKind rowKind = toQuoteRowKind(line.getHeadingLevel());

        for (int index = 0; index < lines.size(); index++) {

            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

            Cell cell = row.createCell(textCol);

            MarkdownRichTextWriter.setResolvedSegmentsCell(ctx.fontCache, cell, lines.get(index), style);

            ctx.st.afterWriteQuotedHeading();

            ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, textCol, rowKind, quoteDepth);
        }
    }

    private static List<List<MdSegment>> parseDisplayLines(String text) {

        List<List<MdSegment>> lines = MarkdownInlineParser.parseParagraphToDisplayLines(text);

        if (!lines.isEmpty()) {
            return lines;
        }

        return Collections.<List<MdSegment>> singletonList(Collections.<MdSegment> emptyList());
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

    private static QuoteRowKind toQuoteRowKind(int headingLevel) {

        switch (headingLevel) {
        case 1:
            return QuoteRowKind.HEADING_1;

        case 2:
            return QuoteRowKind.HEADING_2;

        case 3:
            return QuoteRowKind.HEADING_3;

        default:
            return QuoteRowKind.HEADING_4;
        }
    }
}