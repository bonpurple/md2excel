package md2excel.app;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.util.Arrays;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import md2excel.config.MdFontSettings;
import md2excel.config.MdSheetSettings;

public class MarkdownWorkbookRendererTest {

    @Test
    public void rendersWithoutFileInputOrOutput() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MarkdownWorkbookRenderer renderer = new MarkdownWorkbookRenderer();

            renderer.render(Arrays.asList("## heading", "", "plain paragraph").iterator(), workbook,
                    validFontSettings(), validSheetSettings());

            Sheet sheet = workbook.getSheet(MarkdownWorkbookRenderer.SHEET_NAME);

            assertNotNull(sheet);

            assertEquals("heading", sheet.getRow(1).getCell(1).getStringCellValue());

            /*
             * 見出し後には自動空行が入る。
             */
            assertEquals("plain paragraph", sheet.getRow(3).getCell(1).getStringCellValue());
        }
    }

    @Test
    public void initializesSheetSettings() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MarkdownWorkbookRenderer renderer = new MarkdownWorkbookRenderer();

            renderer.render(Arrays.asList("text").iterator(), workbook, validFontSettings(), validSheetSettings());

            Sheet sheet = workbook.getSheet(MarkdownWorkbookRenderer.SHEET_NAME);

            assertNotNull(sheet);
            assertFalse(sheet.isDisplayGridlines());
            assertFalse(sheet.isPrintGridlines());

            // A列～AN列まで初期化される。
            assertTrue(sheet.getColumnWidth(0) > 0);
            assertTrue(sheet.getColumnWidth(39) > 0);
        }
    }

    @Test
    public void startsRenderingAtB2() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MarkdownWorkbookRenderer renderer = new MarkdownWorkbookRenderer();

            renderer.render(Arrays.asList("plain paragraph").iterator(), workbook, validFontSettings(),
                    validSheetSettings());

            Sheet sheet = workbook.getSheet(MarkdownWorkbookRenderer.SHEET_NAME);

            assertEquals("plain paragraph", sheet.getRow(1).getCell(1).getStringCellValue());

            assertNull(sheet.getRow(1).getCell(2));
        }
    }

    private static MdFontSettings validFontSettings() {
        return new MdFontSettings("Meiryo", 16, 14, 12, 11);
    }

    private static MdSheetSettings validSheetSettings() {
        return new MdSheetSettings(40, VerticalAlignment.BOTTOM);
    }
}