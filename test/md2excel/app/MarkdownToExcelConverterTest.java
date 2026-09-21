package md2excel.app;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import md2excel.config.Md2ExcelConfig;

public class MarkdownToExcelConverterTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void outputStartsAtB2() throws Exception {
        try (XSSFWorkbook workbook = convert("plain paragraph")) {

            Sheet sheet = workbook.getSheet("spec");

            assertCellText(sheet, 1, 1, "plain paragraph");

            assertNull(sheet.getRow(1).getCell(2));
        }
    }

    @Test
    public void pipeInNormalParagraphDoesNotStartTable() throws Exception {

        try (XSSFWorkbook workbook = convert("price | note")) {

            Sheet sheet = workbook.getSheet("spec");

            assertCellText(sheet, 1, 1, "price | note");

            assertNull(sheet.getRow(1).getCell(2));
        }
    }

    @Test
    public void tableWithoutOuterPipesIsRendered() throws Exception {

        String markdown = String.join("\n", "h1 | h2", "--- | ---", "a | b");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            assertCellText(sheet, 1, 1, "h1");
            assertCellText(sheet, 1, 2, "h2");

            assertCellText(sheet, 2, 1, "a");
            assertCellText(sheet, 2, 2, "b");

            assertEquals(BorderStyle.THIN, cell(sheet, 1, 1).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.NONE, cell(sheet, 2, 1).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void escapedPipesRemainInsideTableCells() throws Exception {

        String markdown = String.join("\n", "| h1 | h2 |", "| --- | --- |", "| a \\| b | `x\\|y` |");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            assertCellText(sheet, 2, 1, "a | b");
            assertCellText(sheet, 2, 2, "x|y");
        }
    }

    @Test
    public void tabsInCodeBlockBecomeFourSpacesAndEofClosesFrame() throws Exception {

        // 閉じフェンスなし
        String markdown = String.join("\n", "```text", "\tx", "\ty");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            assertCellText(sheet, 1, 2, "    x");
            assertCellText(sheet, 2, 2, "    y");

            // B2/B3はコードブロック外枠の左端
            assertEquals(BorderStyle.THIN, cell(sheet, 1, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THIN, cell(sheet, 1, 1).getCellStyle().getBorderTop());

            assertEquals(BorderStyle.THIN, cell(sheet, 2, 1).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.THIN, cell(sheet, 2, 1).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void tabsInQuotedCodeBlockBecomeFourSpaces() throws Exception {

        String markdown = String.join("\n", "> ```text", "> \tx", "> \ty", "> ```");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            // 引用内コード本文はD列
            assertCellText(sheet, 1, 3, "    x");
            assertCellText(sheet, 2, 3, "    y");

            // B列は引用罫線、C列はコード枠
            assertEquals(BorderStyle.THICK, cell(sheet, 1, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THIN, cell(sheet, 1, 2).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void nestedQuoteHeadingIsRenderedInDColumn() throws Exception {

        String markdown = String.join("\n", "> > ### nested heading", "> > nested paragraph");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            assertCellText(sheet, 1, 3, "nested heading");

            assertCellText(sheet, 2, 3, "nested paragraph");

            assertEquals(BorderStyle.THICK, cell(sheet, 1, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THICK, cell(sheet, 1, 2).getCellStyle().getBorderLeft());

            Font headingFont = workbook.getFontAt(cell(sheet, 1, 3).getCellStyle().getFontIndex());

            assertTrue(headingFont.getBold());
        }
    }

    @Test
    public void nestedQuoteTableIsRenderedFromDColumn() throws Exception {

        String markdown = String.join("\n", "> > | h1 | h2 |", "> > | --- | --- |", "> > | a | b |");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            assertCellText(sheet, 1, 3, "h1");
            assertCellText(sheet, 1, 4, "h2");

            assertCellText(sheet, 2, 3, "a");
            assertCellText(sheet, 2, 4, "b");

            assertEquals(BorderStyle.THICK, cell(sheet, 1, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THICK, cell(sheet, 1, 2).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THIN, cell(sheet, 1, 3).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.NONE, cell(sheet, 2, 3).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void generatedSheetHasExpectedSettings() throws Exception {

        try (XSSFWorkbook workbook = convert("text")) {
            Sheet sheet = workbook.getSheet("spec");

            assertNotNull(sheet);
            assertFalse(sheet.isDisplayGridlines());
            assertFalse(sheet.isPrintGridlines());

            // A列～AN列まで設定される
            assertTrue(sheet.getColumnWidth(0) > 0);
            assertTrue(sheet.getColumnWidth(39) > 0);
        }
    }

    private XSSFWorkbook convert(String markdown) throws Exception {

        File markdownFile = temporaryFolder.newFile("input.md");

        File excelFile = new File(temporaryFolder.getRoot(), "output.xlsx");

        Files.write(markdownFile.toPath(), markdown.getBytes(StandardCharsets.UTF_8));

        Md2ExcelConfig config = new Md2ExcelConfig(markdownFile.getAbsolutePath(), excelFile.getAbsolutePath(), 40,
                "Meiryo", 16, 14, 12, 11, VerticalAlignment.BOTTOM);

        new MarkdownToExcelConverter().convert(config);

        try (InputStream input = Files.newInputStream(excelFile.toPath())) {

            return new XSSFWorkbook(input);
        }
    }

    private static Cell cell(Sheet sheet, int rowIndex, int columnIndex) {

        Row row = sheet.getRow(rowIndex);
        assertNotNull("Missing row: " + rowIndex, row);

        Cell cell = row.getCell(columnIndex);
        assertNotNull("Missing cell: row=" + rowIndex + ", col=" + columnIndex, cell);

        return cell;
    }

    private static void assertCellText(Sheet sheet, int rowIndex, int columnIndex, String expected) {

        assertEquals(expected, cell(sheet, rowIndex, columnIndex).getStringCellValue());
    }
}