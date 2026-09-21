package md2excel.render;

import java.util.Collections;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

final class BlockQuoteRowUtil {

    private BlockQuoteRowUtil() {
    }

    static Row writeBlankRow(RenderContext ctx, int quoteStartCol, int quoteDepth) {

        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.blankRowStyle);

        int contentCol = RenderLayout.calcQuoteContentCol(quoteStartCol, quoteDepth, ctx.st);

        Cell cell = row.createCell(contentCol);

        MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, Collections.<MarkdownInline.MdSegment>emptyList(),
                ctx.styles.blankRowStyle);

        ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, contentCol, RenderState.QuoteRowKind.BLANK,
                quoteDepth);

        ctx.st.lastRowType = RenderState.RowType.BLANK;
        ctx.st.lastLineWasTable = false;
        ctx.st.lastBlankFromMarkdown = false;
        ctx.st.lastBlankRowIndex = -1;
        ctx.st.lastBlankAfterTable = false;

        ctx.st.lastContentType = RenderState.ContentType.NORMAL;
        ctx.st.lastContentCol = contentCol;
        ctx.st.lastContentWasTable = false;

        ctx.st.lastNormalRowIndex = -1;
        ctx.st.lastNormalIndent = -1;
        ctx.st.bulletDetailActive = false;
        ctx.st.lastWasBlockQuote = true;

        return row;
    }
}