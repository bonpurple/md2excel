package md2excel.app;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**
 * T3: 未変更コードのExcel出力で観察した表の行展開・列配置・罫線を固定する。
 * 空文字の期待値は、文字列セルではなく実際に生成されたBLANKセルを表す。
 */
public class TableRenderingTest {

    @Test
    public void bodyBreaksExpandRowsWithBorderOnlyAtLogicalRowEnd() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "| h1 | h2 |", "| --- | --- |", "| a<br>b | c<br>d |", "| e | f |");

            assertTableSize(sheet, 5, 2);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, "h1", "h2");
            assertRow(workbook, sheet, 3, false, BorderStyle.NONE, "a", "c");
            assertRow(workbook, sheet, 4, false, BorderStyle.HAIR, "b", "d");
            assertRow(workbook, sheet, 5, false, BorderStyle.NONE, "e", "f");
        }
    }

    @Test
    public void unequalBreakCountsPadShorterCellWithStyledBlank() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "| h1 | h2 |", "| --- | --- |", "| a<br>b<br>c | d<br>e |", "| f | g |");

            assertTableSize(sheet, 6, 2);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, "h1", "h2");
            assertRow(workbook, sheet, 3, false, BorderStyle.NONE, "a", "d");
            assertRow(workbook, sheet, 4, false, BorderStyle.NONE, "b", "e");
            assertRow(workbook, sheet, 5, false, BorderStyle.HAIR, "c", "");
            assertRow(workbook, sheet, 6, false, BorderStyle.NONE, "f", "g");
        }
    }

    @Test
    public void trailingBreaksRetainEmptyExpandedRows() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "| h1 | h2 |", "| --- | --- |", "| a<br> | b<br><br> |", "| c | d |");

            assertTableSize(sheet, 6, 2);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, "h1", "h2");
            assertRow(workbook, sheet, 3, false, BorderStyle.NONE, "a", "b");
            assertRow(workbook, sheet, 4, false, BorderStyle.NONE, "", "");
            assertRow(workbook, sheet, 5, false, BorderStyle.HAIR, "", "");
            assertRow(workbook, sheet, 6, false, BorderStyle.NONE, "c", "d");
        }
    }

    @Test
    public void expandedHeaderKeepsBoldAndThinBorderOnEveryRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "| h1<br>h2 | h3 |", "| --- | --- |", "| a | b |");

            assertTableSize(sheet, 4, 2);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, "h1", "h3");
            assertRow(workbook, sheet, 3, true, BorderStyle.THIN, "h2", "");
            assertRow(workbook, sheet, 4, false, BorderStyle.NONE, "a", "b");
        }
    }

    @Test
    public void shortBodyRowsPadToHeaderWidthIncludingFinalRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "| h1 | h2 | h3 |", "| --- | --- | --- |", "| a |", "| b | c | d |",
                    "| e |");

            assertTableSize(sheet, 5, 3);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, "h1", "h2", "h3");
            assertRow(workbook, sheet, 3, false, BorderStyle.HAIR, "a", "", "");
            assertRow(workbook, sheet, 4, false, BorderStyle.HAIR, "b", "c", "d");
            assertRow(workbook, sheet, 5, false, BorderStyle.NONE, "e", "", "");
        }
    }

    @Test
    public void excessBodyCellsAreDiscardedAtHeaderWidth() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "| h1 | h2 |", "| --- | --- |", "| a | b | lost |", "| c | d | lost2 |");

            assertTableSize(sheet, 4, 2);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, "h1", "h2");
            assertRow(workbook, sheet, 3, false, BorderStyle.HAIR, "a", "b");
            assertRow(workbook, sheet, 4, false, BorderStyle.NONE, "c", "d");
        }
    }

    @Test
    public void eofRemovesBottomBorderFromLastExpandedBodyRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "| h1 | h2 |", "| --- | --- |", "| a<br>b | c<br>d |");

            assertTableSize(sheet, 4, 2);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, "h1", "h2");
            assertRow(workbook, sheet, 3, false, BorderStyle.NONE, "a", "c");
            assertRow(workbook, sheet, 4, false, BorderStyle.NONE, "b", "d");
        }
    }

    @Test
    public void quotedExpansionPreservesTableBordersAndDecoratesEveryRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "> | h1<br>h2 | h3 |", "> | --- | --- |", "> | a<br>b | c |",
                    "> | d<br>e | f |");

            assertTableSize(sheet, 7, 38);
            String[][] values = { { "h1", "h3" }, { "h2", "" }, { "a", "c" }, { "b", "" }, { "d", "f" }, { "e", "" } };
            BorderStyle[] borders = { BorderStyle.THIN, BorderStyle.THIN, BorderStyle.NONE, BorderStyle.HAIR,
                    BorderStyle.NONE, BorderStyle.NONE };
            for (int r = 2; r <= 7; r++) {
                Row row = sheet.getRow(r - 1);
                for (int c = 1; c <= 38; c++) {
                    Cell cell = row.getCell(c);
                    boolean tableCell = c == 2 || c == 3;
                    assertValue(cell, tableCell ? values[r - 2][c - 2] : "");
                    assertStyle(workbook, cell, tableCell && r <= 3, tableCell ? borders[r - 2] : BorderStyle.NONE,
                            c == 1 ? BorderStyle.THICK : BorderStyle.NONE, FillPatternType.SOLID_FOREGROUND);
                }
            }
        }
    }

    @Test
    public void threeColumnSettingStillWritesTableThroughD() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 3, "| h1 | h2 | h3 |", "| --- | --- | --- |", "| a | b | c |");

            // Bが描画範囲、Cが右余白、Dが設定範囲外でも切り詰められない。
            assertEquals(3, sheet.getRow(0).getPhysicalNumberOfCells());
            assertTableSize(sheet, 3, 3);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, "h1", "h2", "h3");
            assertRow(workbook, sheet, 3, false, BorderStyle.NONE, "a", "b", "c");
        }
    }

    @Test
    public void twoHundredFiftySixColumnSettingStillWritesTableThroughIW() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            String[] header = new String[256];
            String[] separator = new String[256];
            String[] body = new String[256];
            for (int i = 0; i < 256; i++) {
                header[i] = "h" + (i + 1);
                separator[i] = "---";
                body[i] = "v" + (i + 1);
            }
            Sheet sheet = render(workbook, 256, "|" + String.join("|", header) + "|",
                    "|" + String.join("|", separator) + "|", "|" + String.join("|", body) + "|");

            // IUが描画範囲末尾、IVが右余白、IWが設定範囲外でも全256セルを出力する。
            assertEquals(256, sheet.getRow(0).getPhysicalNumberOfCells());
            assertTableSize(sheet, 3, 256);
            assertRow(workbook, sheet, 2, true, BorderStyle.THIN, header);
            assertRow(workbook, sheet, 3, false, BorderStyle.NONE, body);
        }
    }

    // R1前の実装を実行して観察した値。pipe解除後のinline解析結果も含む。
    @Test
    public void backslashParityDeterminesBodyCellBoundariesAndDisplayedText() throws Exception {
        String[] bodies = { "|a\\|b|tail|", "|a\\\\|b|tail|", "|a\\\\\\|b|tail|", "|a\\\\\\\\|b|tail|" };
        String[][] values = { { "a|b", "tail", "" }, { "a\\", "b", "tail" }, { "a\\|b", "tail", "" },
                { "a\\\\", "b", "tail" } };
        for (int i = 0; i < bodies.length; i++) {
            assertParsedBody(bodies[i], values[i]);
        }
    }

    @Test
    public void unescapedPipeSplitsInlineCodeAndLeavesUnmatchedBackticks() throws Exception {
        assertParsedBody("|`a|b`|tail|", "`a", "b`", "tail");
    }

    @Test
    public void leftOuterPipeAloneIsRemoved() throws Exception {
        assertParsedBody("|a|b", "a", "b", "");
    }

    @Test
    public void rightOuterPipeAloneIsRemoved() throws Exception {
        assertParsedBody("a|b|", "a", "b", "");
    }

    @Test
    public void consecutivePipesCreateMiddleBlankCell() throws Exception {
        assertParsedBody("|a||c|", "a", "", "c");
    }

    @Test
    public void escapedPipeAtLineEndIsRemovedButBackslashRemains() throws Exception {
        assertParsedBody("|a|b\\|", "a", "b\\", "");
    }

    @Test
    public void emptyFirstCellKeepsFollowingTextInSecondColumn() throws Exception {
        assertParsedBody("||b|", "", "b", "");
    }

    @Test
    public void emptyLastCellIsRenderedAsBlank() throws Exception {
        assertParsedBody("|a||", "a", "", "");
    }

    private static void assertParsedBody(String body, String... values) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, 40, "|h1|h2|h3|", "|-|-|-|", body);
            assertTableSize(sheet, 3, 3);
            for (int i = 0; i < values.length; i++) {
                Cell cell = sheet.getRow(2).getCell(i + 1);
                assertNotNull(body + ", column " + (i + 1), cell);
                String label = body + ", " + cell.getAddress();
                assertEquals(label, values[i], cell.getStringCellValue());
                assertEquals(label, values[i].isEmpty() ? CellType.BLANK : CellType.STRING, cell.getCellType());
            }
        }
    }

    private static Sheet render(XSSFWorkbook workbook, int columns, String... lines) {
        return RenderingTestSupport.render(workbook, columns, lines);
    }

    private static void assertTableSize(Sheet sheet, int lastExcelRow, int cellsPerRow) {
        assertEquals(lastExcelRow - 1, sheet.getLastRowNum());
        assertEquals(0, sheet.getNumMergedRegions());
        for (int r = 1; r < lastExcelRow; r++) {
            Row row = sheet.getRow(r);
            assertNotNull("Missing Excel row: " + (r + 1), row);
            assertEquals("Cell count at Excel row: " + (r + 1), cellsPerRow, row.getPhysicalNumberOfCells());
            assertNull(row.getCell(0));
            assertNull(row.getCell(cellsPerRow + 1));
        }
    }

    private static void assertRow(XSSFWorkbook workbook, Sheet sheet, int excelRow, boolean bold, BorderStyle bottom,
            String... values) {
        for (int i = 0; i < values.length; i++) {
            Cell cell = sheet.getRow(excelRow - 1).getCell(i + 1);
            assertValue(cell, values[i]);
            assertStyle(workbook, cell, bold, bottom, BorderStyle.NONE, FillPatternType.NO_FILL);
        }
    }

    private static void assertValue(Cell cell, String value) {
        assertNotNull(cell);
        String address = cell.getAddress().toString();
        assertEquals(address, value.isEmpty() ? CellType.BLANK : CellType.STRING, cell.getCellType());
        assertEquals(address, value, cell.getStringCellValue());
    }

    private static void assertStyle(XSSFWorkbook workbook, Cell cell, boolean bold, BorderStyle bottom,
            BorderStyle left, FillPatternType fill) {
        String address = cell.getAddress().toString();
        CellStyle style = cell.getCellStyle();
        assertEquals(address, left, style.getBorderLeft());
        assertEquals(address, BorderStyle.NONE, style.getBorderTop());
        assertEquals(address, BorderStyle.NONE, style.getBorderRight());
        assertEquals(address, bottom, style.getBorderBottom());
        assertEquals(address, fill, style.getFillPattern());
        assertEquals(address, HorizontalAlignment.GENERAL, style.getAlignment());
        assertEquals(address, VerticalAlignment.BOTTOM, style.getVerticalAlignment());
        assertFalse(address, style.getWrapText());
        Font font = workbook.getFontAt(style.getFontIndex());
        assertEquals(address, 11, font.getFontHeightInPoints());
        assertEquals(address, bold, font.getBold());
    }
}
