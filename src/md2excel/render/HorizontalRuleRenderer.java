package md2excel.render;

import static md2excel.render.RenderLayout.clampCol;

import java.util.Collections;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import md2excel.excel.ExcelCellUtil;

final class HorizontalRuleRenderer {

    private HorizontalRuleRenderer() {
    }

    static void render(RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx, RowUtil.ReuseKind.HORIZONTAL_RULE,
                ctx.styles.blankRowStyle);

        applyHorizontalRuleStyle(row, ctx.styles.horizontalRuleStyle, ctx.st.getStartColIndex(),
                ctx.st.getRenderEndColExclusive());

        ctx.st.afterWriteHorizontalRule();
    }

    private static void applyHorizontalRuleStyle(Row row, CellStyle style, int startCol, int endColExclusive) {

        for (int column = startCol; column < endColExclusive; column++) {

            Cell cell = ExcelCellUtil.getOrCreateCell(row, column);

            cell.setCellStyle(style);
        }
    }

    static void renderQuoted(RenderContext ctx, int quoteStartCol, int quoteDepth) {

        int previousRowIndex = ctx.st.getPreviousRowIndex();

        Row row;

        if (ctx.st.canReusePreviousQuotedBlank()) {
            row = ctx.sheet.getRow(previousRowIndex);

            if (row == null) {
                row = ctx.sheet.createRow(previousRowIndex);
            }
        } else {
            row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.blankRowStyle);

            int contentCol = clampCol(quoteStartCol + quoteDepth - 1, ctx.st);

            Cell cell = row.createCell(contentCol);

            MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell,
                    Collections.<MarkdownInline.MdSegment>emptyList(), ctx.styles.blankRowStyle);
        }

        ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, -1, RenderState.QuoteRowKind.HORIZONTAL_RULE,
                quoteDepth);

        ctx.st.afterWriteQuotedHorizontalRule();
    }
}