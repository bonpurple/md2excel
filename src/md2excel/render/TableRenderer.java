package md2excel.render;

import static md2excel.render.RenderLayout.calcBlockStartCol;
import static md2excel.render.RenderLayout.calcQuoteContentCol;

import md2excel.markdown.LineInfo;

final class TableRenderer {

    private TableRenderer() {
    }

    static void renderSeparator(RenderContext ctx) {
        ctx.st.afterSkipTableSeparatorLine();
    }

    static void renderQuotedSeparator(RenderContext ctx) {
        ctx.st.afterSkipQuotedTableSeparatorLine();
    }

    static void renderRow(LineInfo line, RenderContext ctx) {

        renderTableRow(line.getContentRaw(), calcBlockStartCol(line.getContentIndent(), ctx.st), 0, ctx);
    }

    static void renderQuotedRow(LineInfo line, int quoteStartCol, int quoteDepth, RenderContext ctx) {

        int tableStartCol = calcQuoteContentCol(quoteStartCol, quoteDepth, ctx.st);

        MarkdownTableWriter.TableRowRenderResult result = renderTableRow(line.getContentRaw(), tableStartCol, quoteDepth, ctx);

        for (int rowNum = result.firstRowNum; rowNum <= result.lastRowNum; rowNum++) {

            ctx.st.recordBlockQuoteTableRow(rowNum, quoteStartCol, tableStartCol, result.lastCol, quoteDepth, result.getStyleRole(rowNum));
        }
    }

    private static MarkdownTableWriter.TableRowRenderResult renderTableRow(String tableLine, int firstRowStartCol, int quoteDepth, RenderContext ctx) {

        TableState table = ctx.st.table();

        boolean header = !table.isOpen();

        int tableStartCol = header ? firstRowStartCol : table.getStartCol();

        MarkdownTableWriter.TableRowRenderResult result = MarkdownTableWriter.createTableRows(ctx, tableLine, header, tableStartCol);

        if (header) {
            table.begin(result.firstRowNum, tableStartCol, result.lastCol, quoteDepth);
        } else {
            table.recordBodyRows(result.firstRowNum, result.lastRowNum, result.lastCol);
        }

        if (quoteDepth > 0) {
            ctx.st.afterWriteQuotedTableRow();
        } else {
            ctx.st.afterWriteTableRow();
        }

        return result;
    }
}