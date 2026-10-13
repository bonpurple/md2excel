package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import md2excel.excel.MdStyleCatalog;

public class MarkdownTableColumnLimitTest {

    private static final int MAX_COLUMNS = SheetColumnLayout.EXCEL_MAX_TOTAL_COLUMN_COUNT;

    @Test
    public void headerEndingAtLastPhysicalColumnIsAccepted() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            RenderContext ctx = createContext(workbook);

            MarkdownTableWriter.TableRowRenderResult result = MarkdownTableWriter.createTableRows(ctx, "| left | right |", true, MAX_COLUMNS - 2);

            assertEquals(MAX_COLUMNS - 1, result.lastCol);
            assertEquals("left", ctx.sheet.getRow(1).getCell(MAX_COLUMNS - 2).getStringCellValue());
            assertEquals("right", ctx.sheet.getRow(1).getCell(MAX_COLUMNS - 1).getStringCellValue());
        }
    }

    @Test
    public void headerBeyondLastPhysicalColumnFailsBeforeCreatingRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            RenderContext ctx = createContext(workbook);

            try {
                MarkdownTableWriter.createTableRows(ctx, "| left | right |", true, MAX_COLUMNS - 1);
                fail("IllegalArgumentException was expected");
            } catch (IllegalArgumentException expected) {
                assertNotNull(expected.getMessage());
            }

            assertNull(ctx.sheet.getRow(1));
            assertEquals(0, ctx.sheet.getPhysicalNumberOfRows());
        }
    }

    @Test
    public void excessBodyCellsAreDiscardedBeforeLimitCheck() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            RenderContext ctx = createContext(workbook);
            int lastColumn = MAX_COLUMNS - 1;

            // ヘッダーが1セルの表を模した状態。
            ctx.st.table().begin(1, lastColumn, lastColumn, 0);

            MarkdownTableWriter.TableRowRenderResult result = MarkdownTableWriter.createTableRows(ctx, "| kept | discarded |", false, lastColumn);

            assertEquals(lastColumn, result.lastCol);
            assertEquals("kept", ctx.sheet.getRow(1).getCell(lastColumn).getStringCellValue());
            assertEquals(1, ctx.sheet.getRow(1).getPhysicalNumberOfCells());
        }
    }

    private static RenderContext createContext(XSSFWorkbook workbook) {
        Sheet sheet = workbook.createSheet("spec");

        MdStyleCatalog styles = new MdStyleCatalog(workbook, "Meiryo", 16, 14, 12, 11, VerticalAlignment.BOTTOM);

        // 設定上の描画可能列はB列だけ。
        // それでも表にはExcelの物理上限までの列を許可する。
        SheetColumnLayout layout = new SheetColumnLayout(3, 1, 1);

        return new RenderContext(workbook, sheet, styles, layout, 1);
    }
}