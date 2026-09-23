package md2excel.app;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import md2excel.config.MdFontSettings;
import md2excel.config.MdSheetSettings;

/**
 * 未変更の描画処理で観察した段落・Setextの出力を固定する。 一般的なMarkdown仕様への適合性ではなく、現在のセル値・位置・書式を検証する。
 */
public class ParagraphRenderingTest {

    @Test
    public void joinsNormalParagraphLinesWithSpace() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first", "second");

            assertTextCells(sheet, "B2=first second");
            assertEquals(1, sheet.getLastRowNum());
        }
    }

    @Test
    public void trailingTwoSpacesCreateNextExcelRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first  ", "second");

            assertTextCells(sheet, "B2=first", "B3=second");
            assertEquals(2, sheet.getLastRowNum());
        }
    }

    @Test
    public void trailingBackslashCreatesNextExcelRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first\\", "second");

            assertTextCells(sheet, "B2=first", "B3=second");
            assertEquals(2, sheet.getLastRowNum());
        }
    }

    @Test
    public void brCreatesNextExcelRow() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first<br>second");

            assertTextCells(sheet, "B2=first", "B3=second");
            assertEquals(2, sheet.getLastRowNum());
        }
    }

    @Test
    public void boldSpansInputLines() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "**first", "second**");

            assertTextCells(sheet, "B2=first second");
            assertEquals(1, sheet.getLastRowNum());
            assertRichTextFont(workbook, sheet.getRow(1).getCell(1), true, false);
        }
    }

    @Test
    public void italicSpansInputLines() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "*first", "second*");

            assertTextCells(sheet, "B2=first second");
            assertEquals(1, sheet.getLastRowNum());
            assertRichTextFont(workbook, sheet.getRow(1).getCell(1), false, true);
        }
    }

    @Test
    public void bulletContinuationJoinsFirstCell() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "- first", "  second");

            assertTextCells(sheet, "C2=・ first second");
            assertEquals(1, sheet.getLastRowNum());
        }
    }

    @Test
    public void numberedContinuationJoinsFirstCell() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "1. first", "   second");

            assertTextCells(sheet, "C2=1. first second");
            assertEquals(1, sheet.getLastRowNum());
        }
    }

    @Test
    public void bulletHardBreakMovesContentToContinuationColumnWithoutMarker() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "- first  ", "  second");

            assertTextCells(sheet, "C2=・ first", "D3=second");
            assertEquals(2, sheet.getLastRowNum());
            assertListLineStyles(workbook, sheet.getRow(1).getCell(2), sheet.getRow(2).getCell(3));
        }
    }

    @Test
    public void numberedHardBreakPreservesParenthesizedMarkerOnlyOnFirstLine() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "12) first  ", "    second");

            assertTextCells(sheet, "C2=12) first", "D3=second");
            assertEquals(2, sheet.getLastRowNum());
            assertListLineStyles(workbook, sheet.getRow(1).getCell(2), sheet.getRow(2).getCell(3));
        }
    }

    @Test
    public void bulletHardBreakUsesOnlyRenderableColumnAtMinimumColumnCount() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, MdSheetSettings.MIN_TOTAL_COLUMN_COUNT, "- first  ", "  second");

            assertTextCells(sheet, "B2=・ first", "B3=second");
            assertEquals(2, sheet.getLastRowNum());
            assertListLineStyles(workbook, sheet.getRow(1).getCell(1), sheet.getRow(2).getCell(1));
        }
    }

    @Test
    public void numberedHardBreakUsesOnlyRenderableColumnAtMinimumColumnCount() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, MdSheetSettings.MIN_TOTAL_COLUMN_COUNT, "12) first  ", "    second");

            assertTextCells(sheet, "B2=12) first", "B3=second");
            assertEquals(2, sheet.getLastRowNum());
            assertListLineStyles(workbook, sheet.getRow(1).getCell(1), sheet.getRow(2).getCell(1));
        }
    }

    @Test
    public void quotedParagraphLinesJoinAtC2() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "> first", "> second");

            assertTextCells(sheet, "C2=first second");
            assertEquals(1, sheet.getLastRowNum());
            assertEquals(BorderStyle.THICK, sheet.getRow(1).getCell(1).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void quotedBulletAndNumberKeepMarkersAndJoinedContinuations() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "> - bullet", ">   continued", "> 12) number", ">     continued");

            assertTextCells(sheet, "D2=・ bullet continued", "D3=12) number continued");
            assertEquals(2, sheet.getLastRowNum());
            assertEquals(BorderStyle.THICK, sheet.getRow(1).getCell(1).getCellStyle().getBorderLeft());
            assertEquals(BorderStyle.THICK, sheet.getRow(2).getCell(1).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void nestedQuotedListsKeepDepthAndHardBreakContinuationColumns() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, ">> - outer  ", ">>   continued", ">>   2. nested  ",
                    ">>      continued");

            assertTextCells(sheet, "E2=・ outer", "F3=continued", "F4=2. nested", "G5=continued");
            assertEquals(4, sheet.getLastRowNum());
            assertEquals(BorderStyle.THICK, sheet.getRow(1).getCell(1).getCellStyle().getBorderLeft());
            assertEquals(BorderStyle.THICK, sheet.getRow(4).getCell(1).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void equalsUnderlineCreatesLevelOneHeading() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "title", "===");

            assertTextCells(sheet, "B2=title");
            assertEquals(1, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(1), 16, true);
        }
    }

    @Test
    public void dashUnderlineCreatesLevelTwoHeading() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "title", "---");

            assertTextCells(sheet, "B2=title");
            assertEquals(1, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(1), 14, true);
            assertEquals(BorderStyle.NONE, sheet.getRow(1).getCell(1).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void blankBeforeDashesProducesRuleInsteadOfHeading() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "title", "", "---");

            assertTextCells(sheet, "B2=title");
            assertEquals(2, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(1), 11, false);
            assertEquals(BorderStyle.HAIR, sheet.getRow(2).getCell(1).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void quotedEqualsUnderlineRemainsParagraphText() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "> title", "> ===");

            assertTextCells(sheet, "C2=title ===");
            assertEquals(1, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(2), 11, false);
        }
    }

    @Test
    public void quotedDashUnderlineProducesRule() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "> title", "> ---");

            assertTextCells(sheet, "C2=title");
            assertEquals(2, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(2), 11, false);
            assertEquals(BorderStyle.THICK, sheet.getRow(2).getCell(1).getCellStyle().getBorderLeft());
            assertEquals(BorderStyle.HAIR, sheet.getRow(2).getCell(2).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void bulletEqualsUnderlineRemainsContinuationText() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "- title", "  ===");

            assertTextCells(sheet, "C2=・ title ===");
            assertEquals(1, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(2), 11, false);
        }
    }

    @Test
    public void bulletDashUnderlineProducesRule() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "- title", "  ---");

            assertTextCells(sheet, "C2=・ title");
            assertEquals(2, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(2), 11, false);
            assertEquals(BorderStyle.HAIR, sheet.getRow(2).getCell(1).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void numberedEqualsUnderlineRemainsContinuationText() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "1. title", "   ===");

            assertTextCells(sheet, "C2=1. title ===");
            assertEquals(1, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(2), 11, false);
        }
    }

    @Test
    public void numberedDashUnderlineProducesRule() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "1. title", "   ---");

            assertTextCells(sheet, "C2=1. title");
            assertEquals(2, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(2), 11, false);
            assertEquals(BorderStyle.HAIR, sheet.getRow(2).getCell(1).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void setextHeadingIncludesAllJoinedParagraphLines() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = render(workbook, "first", "second", "===");

            assertTextCells(sheet, "B2=first second");
            assertEquals(1, sheet.getLastRowNum());
            assertCellFont(workbook, sheet.getRow(1).getCell(1), 16, true);
        }
    }

    private static Sheet render(XSSFWorkbook workbook, String... lines) {
        return render(workbook, 40, lines);
    }

    private static Sheet render(XSSFWorkbook workbook, int totalColumnCount, String... lines) {
        new MarkdownWorkbookRenderer().render(Arrays.asList(lines).iterator(), workbook,
                new MdFontSettings("Meiryo", 16, 14, 12, 11),
                new MdSheetSettings(totalColumnCount, VerticalAlignment.BOTTOM));
        return workbook.getSheet("spec");
    }

    private static void assertTextCells(Sheet sheet, String... expected) {
        List<String> actual = new ArrayList<String>();

        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && !cell.getStringCellValue().isEmpty()) {
                    actual.add(cell.getAddress().toString() + "=" + cell.getStringCellValue());
                }
            }
        }

        assertEquals(Arrays.asList(expected), actual);
    }

    private static void assertCellFont(XSSFWorkbook workbook, Cell cell, int size, boolean bold) {
        Font font = workbook.getFontAt(cell.getCellStyle().getFontIndex());

        assertEquals(size, font.getFontHeightInPoints());
        assertEquals(bold, font.getBold());
    }

    private static void assertListLineStyles(XSSFWorkbook workbook, Cell first, Cell continuation) {
        assertPlainCell(workbook, first);
        assertPlainCell(workbook, continuation);
        assertEquals(first.getCellStyle().getIndex(), continuation.getCellStyle().getIndex());
    }

    private static void assertPlainCell(XSSFWorkbook workbook, Cell cell) {
        CellStyle style = cell.getCellStyle();

        assertEquals(BorderStyle.NONE, style.getBorderLeft());
        assertEquals(BorderStyle.NONE, style.getBorderTop());
        assertEquals(BorderStyle.NONE, style.getBorderRight());
        assertEquals(BorderStyle.NONE, style.getBorderBottom());
        assertEquals(FillPatternType.NO_FILL, style.getFillPattern());
        assertEquals(VerticalAlignment.BOTTOM, style.getVerticalAlignment());
        assertFalse(style.getWrapText());
        assertCellFont(workbook, cell, 11, false);
    }

    private static void assertRichTextFont(XSSFWorkbook workbook, Cell cell, boolean bold, boolean italic) {
        XSSFRichTextString rich = (XSSFRichTextString) cell.getRichStringCellValue();

        // POIの書式runの分割数ではなく、各文字に適用される書式を確認する。
        for (int index = 0; index < rich.length(); index++) {
            XSSFFont font = rich.getFontAtIndex(index);
            if (font == null) {
                font = workbook.getFontAt(cell.getCellStyle().getFontIndex());
            }

            assertEquals("bold at character " + index, bold, font.getBold());
            assertEquals("italic at character " + index, italic, font.getItalic());
        }
    }
}
