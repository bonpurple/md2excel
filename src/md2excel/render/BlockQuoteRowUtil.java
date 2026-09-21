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

        ctx.st.afterWriteQuotedBlank(row.getRowNum(), quoteStartCol, contentCol, quoteDepth);

        return row;
    }
}