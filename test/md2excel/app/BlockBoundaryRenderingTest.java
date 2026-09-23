package md2excel.app;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import java.util.Arrays;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**
 * T2: 未変更コードで観察した空行・ブロック境界・EOFの描画結果を固定する。
 * Markdown仕様から期待値を推測せず、セル位置・値・罫線・書式を検証する。
 */
public class BlockBoundaryRenderingTest {

    @Test
    public void headingFollowedDirectlyByParagraphInsertsBlankRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "# heading", "body");

            assertOutput(sheet, 4, "B2=heading", "B4=body");
            assertEmptyRow(sheet, 3);
            assertFont(workbook, cell(sheet, "B2"), 16, true);
            assertPlainCell(workbook, cell(sheet, "B4"));
        }
    }

    @Test
    public void headingFollowedByBulletInsertsBlankRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "# heading", "- item");

            assertOutput(sheet, 4, "B2=heading", "C4=・ item");
            assertEmptyRow(sheet, 3);
            assertPlainCell(workbook, cell(sheet, "C4"));
        }
    }

    @Test
    public void headingFollowedByNumberedItemInsertsBlankRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "# heading", "1. item");

            assertOutput(sheet, 4, "B2=heading", "C4=1. item");
            assertEmptyRow(sheet, 3);
            assertPlainCell(workbook, cell(sheet, "C4"));
        }
    }

    @Test
    public void tableBlankThenParagraphKeepsSeparatorAndFinalizesTable() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "| h1 | h2 |", "| --- | --- |", "| a | b |", "", "body");

            assertOutput(sheet, 5, "B2=h1", "C2=h2", "B3=a", "C3=b", "B5=body");
            assertTableEnd(workbook, sheet);
            assertEmptyRow(sheet, 4);
            assertPlainCell(workbook, cell(sheet, "B5"));
        }
    }

    @Test
    public void tableBlankThenCodeKeepsSeparatorAndClosesFrame() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "| h1 | h2 |", "| --- | --- |", "| a | b |", "", "```", "code", "```");

            assertOutput(sheet, 5, "B2=h1", "C2=h2", "B3=a", "C3=b", "C5=code");
            assertTableEnd(workbook, sheet);
            assertEmptyRow(sheet, 4);
            assertSingleRowCodeFrame(workbook, sheet, 5);
        }
    }

    @Test
    public void quoteThenParagraphRemovesDecorationWithoutBlankRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "> quote", "body");

            assertOutput(sheet, 3, "C2=quote", "B3=body");
            assertEquals(BorderStyle.THICK, cell(sheet, "B2").getCellStyle().getBorderLeft());
            assertEquals(FillPatternType.SOLID_FOREGROUND, cell(sheet, "C2").getCellStyle().getFillPattern());
            assertEquals(FillPatternType.SOLID_FOREGROUND, cell(sheet, "AM2").getCellStyle().getFillPattern());
            assertPlainCell(workbook, cell(sheet, "B3"));
            assertEquals(1, sheet.getRow(2).getPhysicalNumberOfCells());
        }
    }

    @Test
    public void returningFromDeepListInsertsBlankAtEachShallowerLevel() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "- root", "  - child", "    - deep", "  - shallow", "- sibling");

            assertOutput(sheet, 8, "C2=・ root", "D3=・ child", "E4=・ deep", "D6=・ shallow", "C8=・ sibling");
            assertEmptyRow(sheet, 5);
            assertEmptyRow(sheet, 7);
        }
    }

    @Test
    public void mixedNestedListKeepsIndentedParagraphAndBlankBeforeShallowReturn() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "- root", "  1. child", "", "    detail", "- sibling");

            assertOutput(sheet, 6, "C2=・ root", "D3=1. child", "E4=detail", "C6=・ sibling");
            assertPlainCell(workbook, cell(sheet, "C2"));
            assertPlainCell(workbook, cell(sheet, "D3"));
            assertPlainCell(workbook, cell(sheet, "E4"));
            assertEmptyRow(sheet, 5);
            assertPlainCell(workbook, cell(sheet, "C6"));
        }
    }

    @Test
    public void consecutiveBlanksBetweenParagraphsProduceOneBlankRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first", "", "", "", "second");

            assertOutput(sheet, 4, "B2=first", "B4=second");
            assertEmptyRow(sheet, 3);
        }
    }

    @Test
    public void consecutiveBlanksAtEofLeaveOneBlankRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first", "", "", "");

            assertOutput(sheet, 3, "B2=first");
            assertEmptyRow(sheet, 3);
        }
    }

    @Test
    public void horizontalRuleReusesPrecedingBlankAndConsumesFollowingBlank() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first", "", "---", "", "second");

            assertOutput(sheet, 4, "B2=first", "B4=second");
            for (int column = 1; column <= 38; column++) {
                Cell rule = sheet.getRow(2).getCell(column);
                assertNotNull(rule);
                assertEquals(BorderStyle.HAIR, rule.getCellStyle().getBorderBottom());
            }
            assertNull(sheet.getRow(2).getCell(0));
            assertNull(sheet.getRow(2).getCell(39));
            assertPlainCell(workbook, cell(sheet, "B2"));
            assertPlainCell(workbook, cell(sheet, "B4"));
        }
    }

    @Test
    public void closedEmptyCodeBlockProducesNoRowsOrFrame() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "```", "```");

            assertOutput(sheet, 1);
            assertNull(sheet.getRow(1));
        }
    }

    @Test
    public void emptyCodeBlockSeparatesParagraphsWithoutAddingBlankRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first", "```", "```", "second");

            assertOutput(sheet, 3, "B2=first", "B3=second");
            assertPlainCell(workbook, cell(sheet, "B2"));
            assertPlainCell(workbook, cell(sheet, "B3"));
            assertEquals(1, sheet.getRow(1).getPhysicalNumberOfCells());
            assertEquals(1, sheet.getRow(2).getPhysicalNumberOfCells());
        }
    }

    @Test
    public void unclosedEmptyCodeBlockAtEofProducesNoRowsOrFrame() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "```");

            assertOutput(sheet, 1);
            assertNull(sheet.getRow(1));
        }
    }

    @Test
    public void codeBlockContainingOneEmptyLineStillProducesFrame() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "```", "", "```");

            assertOutput(sheet, 2);
            assertEquals("", cell(sheet, "C2").getStringCellValue());
            assertSingleRowCodeFrame(workbook, sheet, 2);
        }
    }

    @Test
    public void quotedTableAtEofFinalizesLastBodyBorderBeforeQuoteDecoration() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "> | h1 | h2 |", "> | --- | --- |", "> | a | b |", "> | c | d |");

            assertOutput(sheet, 4, "C2=h1", "D2=h2", "C3=a", "D3=b", "C4=c", "D4=d");
            for (int row = 2; row <= 4; row++) {
                assertEquals(BorderStyle.THICK, cell(sheet, "B" + row).getCellStyle().getBorderLeft());
                for (String column : Arrays.asList("C", "D")) {
                    Cell tableCell = cell(sheet, column + row);
                    assertEquals(row == 2 ? BorderStyle.THIN : row == 3 ? BorderStyle.HAIR : BorderStyle.NONE,
                            tableCell.getCellStyle().getBorderBottom());
                    assertEquals(FillPatternType.SOLID_FOREGROUND, tableCell.getCellStyle().getFillPattern());
                    assertFont(workbook, tableCell, 11, row == 2);
                }
                assertEquals(BorderStyle.NONE, cell(sheet, "AM" + row).getCellStyle().getBorderBottom());
                assertEquals(BorderStyle.NONE, cell(sheet, "AM" + row).getCellStyle().getBorderLeft());
                assertEquals(BorderStyle.NONE, cell(sheet, "AM" + row).getCellStyle().getBorderRight());
                assertEquals(BorderStyle.NONE, cell(sheet, "AM" + row).getCellStyle().getBorderTop());
                assertEquals(FillPatternType.SOLID_FOREGROUND, cell(sheet, "AM" + row).getCellStyle().getFillPattern());
            }
        }
    }

    private static Sheet render(XSSFWorkbook workbook, String... lines) {
        return RenderingTestSupport.render(workbook, lines);
    }

    private static void assertOutput(Sheet sheet, int lastExcelRow, String... expected) {
        assertEquals(Arrays.asList(expected), RenderingTestSupport.textCells(sheet));
        assertEquals(lastExcelRow - 1, sheet.getLastRowNum());
    }

    private static Cell cell(Sheet sheet, String address) {
        CellReference reference = new CellReference(address);
        Row row = sheet.getRow(reference.getRow());
        assertNotNull("Missing row: " + address, row);
        Cell cell = row.getCell(reference.getCol());
        assertNotNull("Missing cell: " + address, cell);
        return cell;
    }

    private static void assertEmptyRow(Sheet sheet, int excelRow) {
        Row row = sheet.getRow(excelRow - 1);
        assertNotNull("Missing blank row: " + excelRow, row);
        assertEquals(0, row.getPhysicalNumberOfCells());
    }

    private static void assertFont(XSSFWorkbook workbook, Cell cell, int size, boolean bold) {
        Font font = workbook.getFontAt(cell.getCellStyle().getFontIndex());
        assertEquals(size, font.getFontHeightInPoints());
        assertEquals(bold, font.getBold());
    }

    private static void assertPlainCell(XSSFWorkbook workbook, Cell cell) {
        CellStyle style = cell.getCellStyle();
        assertEquals(BorderStyle.NONE, style.getBorderLeft());
        assertEquals(BorderStyle.NONE, style.getBorderTop());
        assertEquals(BorderStyle.NONE, style.getBorderRight());
        assertEquals(BorderStyle.NONE, style.getBorderBottom());
        assertEquals(FillPatternType.NO_FILL, style.getFillPattern());
        assertFont(workbook, cell, 11, false);
    }

    private static void assertTableEnd(XSSFWorkbook workbook, Sheet sheet) {
        for (String column : Arrays.asList("B", "C")) {
            assertEquals(BorderStyle.THIN, cell(sheet, column + "2").getCellStyle().getBorderBottom());
            assertFont(workbook, cell(sheet, column + "2"), 11, true);
            assertPlainCell(workbook, cell(sheet, column + "3"));
        }
    }

    private static void assertSingleRowCodeFrame(XSSFWorkbook workbook, Sheet sheet, int excelRow) {
        Row row = sheet.getRow(excelRow - 1);
        assertEquals(38, row.getPhysicalNumberOfCells());
        assertNull(row.getCell(0));
        assertNull(row.getCell(39));
        for (int column = 1; column <= 38; column++) {
            Cell cell = row.getCell(column);
            assertNotNull(cell);
            CellStyle style = cell.getCellStyle();
            assertEquals(BorderStyle.THIN, style.getBorderTop());
            assertEquals(BorderStyle.THIN, style.getBorderBottom());
            assertEquals(column == 1 ? BorderStyle.THIN : BorderStyle.NONE, style.getBorderLeft());
            assertEquals(column == 38 ? BorderStyle.THIN : BorderStyle.NONE, style.getBorderRight());
            assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
            assertFont(workbook, cell, 10, false);
        }
    }
}
