package md2excel.render;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import md2excel.excel.MdStyleCatalog;
import md2excel.markdown.LineInfo;

public class MdBlockBoundaryTest {

    @Test
    public void continuingTableKeepsTableOpen() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            RenderContext ctx = createOpenTableContext(workbook);
            LineInfo line = MarkdownLineParser.parse("| body |", ctx.st, true);

            MdBlockBoundary.processLineBoundary(line, ctx);

            assertTrue(ctx.st.isLastLineTable());
            assertTrue(ctx.st.table().isOpen());
        }
    }

    @Test
    public void leavingTableForQuotedParagraphClosesTable() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            RenderContext ctx = createOpenTableContext(workbook);
            LineInfo line = MarkdownLineParser.parse("> paragraph", ctx.st, false);

            MdBlockBoundary.processLineBoundary(line, ctx);

            assertFalse(ctx.st.isLastLineTable());
            assertFalse(ctx.st.table().isOpen());
        }
    }

    @Test
    public void directApplyStillClosesTableWhenPolicyRequestsIt() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            RenderContext ctx = createOpenTableContext(workbook);

            MdBlockBoundary.apply(MdBlockBoundary.Policy.HEADING, ctx);

            assertFalse(ctx.st.isLastLineTable());
            assertFalse(ctx.st.table().isOpen());
        }
    }

    private static RenderContext createOpenTableContext(XSSFWorkbook workbook) {
        MdStyleCatalog styles = new MdStyleCatalog(workbook, "Meiryo", 16, 14, 12, 11, VerticalAlignment.BOTTOM);

        RenderContext ctx = new RenderContext(workbook, workbook.createSheet("spec"), styles, new SheetColumnLayout(40, 1, 1), 1);

        // 本文行のない、ヘッダーだけの表として状態を用意する。
        ctx.st.table().begin(1, 1, 1, 0);
        ctx.st.afterWriteTableRow();

        return ctx;
    }
}