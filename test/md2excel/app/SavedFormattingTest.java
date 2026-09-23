package md2excel.app;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import md2excel.config.MdFontSettings;
import md2excel.config.MdSheetSettings;

/**
 * T4: 未変更コードをxlsxへ保存・再読込して観察した文字書式・セル書式を固定する。
 * 期待値の区間は読みやすさのためのもので、POIのrun分割とは対応させない。
 */
public class SavedFormattingTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void boldEndsBeforeFollowingPlainText() throws Exception {
        try (XSSFWorkbook workbook = reopen("p **太A** q")) {
            assertText(workbook, "B2", plain("p "), text("太A", true, false), plain(" q"));
        }
    }

    @Test
    public void italicEndsBeforeFollowingPlainText() throws Exception {
        try (XSSFWorkbook workbook = reopen("p *斜A* q")) {
            assertText(workbook, "B2", plain("p "), text("斜A", false, true), plain(" q"));
        }
    }

    @Test
    public void boldItalicEndsBeforeFollowingPlainText() throws Exception {
        try (XSSFWorkbook workbook = reopen("p ***両A*** q")) {
            assertText(workbook, "B2", plain("p "), text("両A", true, true), plain(" q"));
        }
    }

    @Test
    public void headingEmphasisPreservesBaseBoldAndSize() throws Exception {
        try (XSSFWorkbook workbook = reopen("# H *斜* **太** ***両*** `A日`")) {
            assertText(workbook, "B2", font("H ", "Meiryo", 16, true, false), font("斜", "Meiryo", 16, true, true),
                    font(" 太 ", "Meiryo", 16, true, false), font("両", "Meiryo", 16, true, true),
                    font(" ", "Meiryo", 16, true, false), code("A", "Consolas", 16, true),
                    code("日", "Meiryo", 16, true));
        }
    }

    @Test
    public void inlineCodeKeepsLiteralMarkersAndSwitchesAsciiAndJapaneseFonts() throws Exception {
        try (XSSFWorkbook workbook = reopen("p `*A日*` q")) {
            assertText(workbook, "B2", plain("p "), code("*A", "Consolas", 11, false), code("日", "Meiryo", 11, false),
                    code("*", "Consolas", 11, false), plain(" q"));
            assertEquals(FillPatternType.NO_FILL, cell(workbook, "B2").getCellStyle().getFillPattern());
        }
    }

    @Test
    public void codeInsideBoldItalicKeepsBoldButDropsItalic() throws Exception {
        try (XSSFWorkbook workbook = reopen("***`A日`***")) {
            assertText(workbook, "B2", code("A", "Consolas", 11, true), code("日", "Meiryo", 11, true));
        }
    }

    @Test
    public void plainAsciiAndJapaneseUseSameFont() throws Exception {
        try (XSSFWorkbook workbook = reopen("A日本B")) {
            assertText(workbook, "B2", plain("A日本B"));
        }
    }

    @Test
    public void boldAcrossSoftBreakIncludesInsertedSpace() throws Exception {
        try (XSSFWorkbook workbook = reopen("p **A", "日** q")) {
            assertText(workbook, "B2", plain("p "), text("A 日", true, false), plain(" q"));
            assertEquals(1, workbook.getSheet("spec").getLastRowNum());
        }
    }

    @Test
    public void boldItalicAcrossHardBreakContinuesInNextCell() throws Exception {
        try (XSSFWorkbook workbook = reopen("p ***A  ", "日*** q")) {
            assertText(workbook, "B2", plain("p "), text("A", true, true));
            assertText(workbook, "B3", text("日", true, true), plain(" q"));
        }
    }

    @Test
    public void italicAcrossBrContinuesInNextCell() throws Exception {
        try (XSSFWorkbook workbook = reopen("p *A<br>日* q")) {
            assertText(workbook, "B2", plain("p "), text("A", false, true));
            assertText(workbook, "B3", text("日", false, true), plain(" q"));
        }
    }

    @Test
    public void hardBreakInsideCodeBecomesCodeStyledSpace() throws Exception {
        try (XSSFWorkbook workbook = reopen("p `A  ", "日` q")) {
            assertText(workbook, "B2", plain("p "), code("A ", "Consolas", 11, false), code("日", "Meiryo", 11, false),
                    plain(" q"));
            assertEquals(1, workbook.getSheet("spec").getLastRowNum());
        }
    }

    @Test
    public void quoteBackgroundCoexistsWithRichText() throws Exception {
        try (XSSFWorkbook workbook = reopen("> p **A日** q")) {
            assertText(workbook, "C2", plain("p "), text("A日", true, false), plain(" q"));
            assertGrayFill(cell(workbook, "C2").getCellStyle());
            assertGrayFill(cell(workbook, "B2").getCellStyle());
            assertEquals(BorderStyle.THICK, cell(workbook, "B2").getCellStyle().getBorderLeft());
            assertArrayEquals(new byte[] { 0, 112, (byte) 192 },
                    cell(workbook, "B2").getCellStyle().getLeftBorderXSSFColor().getRGB());
        }
    }

    @Test
    public void tableHeaderRichTextInheritsBoldAlongsideThinBorder() throws Exception {
        try (XSSFWorkbook workbook = table()) {
            assertText(workbook, "B2", text("H ", true, false), text("斜", true, true), text(" ", true, false),
                    code("A", "Consolas", 11, true), code("日", "Meiryo", 11, true));
            assertEquals(BorderStyle.THIN, cell(workbook, "B2").getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void tableBodyRichTextKeepsPlainBaseAndFinalRowBorder() throws Exception {
        try (XSSFWorkbook workbook = table()) {
            assertText(workbook, "B3", plain("p "), text("太", true, false), plain(" "),
                    code("A", "Consolas", 11, false), code("日", "Meiryo", 11, false));
            assertEquals(BorderStyle.NONE, cell(workbook, "B3").getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void codeBlockUsesTenPointAsciiAndJapaneseFonts() throws Exception {
        try (XSSFWorkbook workbook = reopen("```", "A日", "B本", "```")) {
            assertText(workbook, "C2", font("A", "Consolas", 10, false, false), font("日", "Meiryo", 10, false, false));
            assertText(workbook, "C3", font("B", "Consolas", 10, false, false), font("本", "Meiryo", 10, false, false));
        }
    }

    @Test
    public void codeBlockSwitchesBackToAsciiFontAfterJapanese() throws Exception {
        try (XSSFWorkbook workbook = reopen("```", "A日B", "```")) {
            assertText(workbook, "C2", font("A", "Consolas", 10, false, false),
                    font("日", "Meiryo", 10, false, false), font("B", "Consolas", 10, false, false));
        }
    }

    @Test
    public void codeBlockGrayFillAndFrameCoverBothRows() throws Exception {
        try (XSSFWorkbook workbook = reopen("```", "A日", "B本", "```")) {
            for (int row = 1; row <= 2; row++) {
                for (int column = 1; column <= 38; column++) {
                    XSSFCellStyle style = workbook.getSheet("spec").getRow(row).getCell(column).getCellStyle();
                    assertGrayFill(style);
                    assertEquals(row == 1 ? BorderStyle.THIN : BorderStyle.NONE, style.getBorderTop());
                    assertEquals(row == 2 ? BorderStyle.THIN : BorderStyle.NONE, style.getBorderBottom());
                    assertEquals(column == 1 ? BorderStyle.THIN : BorderStyle.NONE, style.getBorderLeft());
                    assertEquals(column == 38 ? BorderStyle.THIN : BorderStyle.NONE, style.getBorderRight());
                    assertAlignment(style, VerticalAlignment.BOTTOM);
                }
            }
        }
    }

    @Test
    public void customFontAndBodySizeApplyToEmphasisButCodeKeepsOwnNames() throws Exception {
        try (XSSFWorkbook workbook = reopen(customFonts(), VerticalAlignment.BOTTOM, "p ***日*** `A日`")) {
            assertText(workbook, "B2", font("p ", "Arial", 13, false, false), font("日", "Arial", 13, true, true),
                    font(" ", "Arial", 13, false, false), code("A", "Consolas", 13, false),
                    code("日", "Meiryo", 13, false));
        }
    }

    @Test
    public void customHeadingSizesApplyToEveryCharacterIncludingEmphasis() throws Exception {
        try (XSSFWorkbook workbook = reopen(customFonts(), VerticalAlignment.BOTTOM, "# H *日*", "", "## H *日*", "",
                "### H *日*", "", "#### H *日*")) {
            String[] addresses = { "B2", "B4", "B6", "B8" };
            int[] sizes = { 23, 19, 17, 13 };
            for (int i = 0; i < addresses.length; i++) {
                assertText(workbook, addresses[i], font("H ", "Arial", sizes[i], true, false),
                        font("日", "Arial", sizes[i], true, true));
            }
        }
    }

    @Test
    public void customFontSettingsDoNotChangeCodeBlockFontsOrSize() throws Exception {
        try (XSSFWorkbook workbook = reopen(customFonts(), VerticalAlignment.BOTTOM, "```", "A日", "```")) {
            assertText(workbook, "C2", font("A", "Consolas", 10, false, false), font("日", "Meiryo", 10, false, false));
        }
    }

    @Test
    public void topAlignmentSurvivesForParagraphAndHeading() throws Exception {
        try (XSSFWorkbook workbook = reopen(customFonts(), VerticalAlignment.TOP, "# H *日*", "", "p ***日*** `A日`")) {
            assertAlignment(cell(workbook, "B2").getCellStyle(), VerticalAlignment.TOP);
            assertAlignment(cell(workbook, "B4").getCellStyle(), VerticalAlignment.TOP);
        }
    }

    @Test
    public void topAlignmentSurvivesQuoteDecoration() throws Exception {
        try (XSSFWorkbook workbook = reopen(customFonts(), VerticalAlignment.TOP, "> p **A日** q")) {
            assertAlignment(cell(workbook, "B2").getCellStyle(), VerticalAlignment.TOP);
            assertAlignment(cell(workbook, "C2").getCellStyle(), VerticalAlignment.TOP);
        }
    }

    @Test
    public void topAlignmentSurvivesTableStyles() throws Exception {
        try (XSSFWorkbook workbook = reopen(customFonts(), VerticalAlignment.TOP, "| H *斜* `A日` |", "| --- |",
                "| p **太** `A日` |")) {
            assertAlignment(cell(workbook, "B2").getCellStyle(), VerticalAlignment.TOP);
            assertAlignment(cell(workbook, "B3").getCellStyle(), VerticalAlignment.TOP);
        }
    }

    @Test
    public void topAlignmentSurvivesCodeFrameStyles() throws Exception {
        try (XSSFWorkbook workbook = reopen(customFonts(), VerticalAlignment.TOP, "```", "A日", "```")) {
            for (int column = 1; column <= 38; column++) {
                assertAlignment(workbook.getSheet("spec").getRow(1).getCell(column).getCellStyle(),
                        VerticalAlignment.TOP);
            }
        }
    }

    private XSSFWorkbook table() throws Exception {
        return reopen("| H *斜* `A日` |", "| --- |", "| p **太** `A日` |");
    }

    private XSSFWorkbook reopen(String... lines) throws Exception {
        return reopen(new MdFontSettings("Meiryo", 16, 14, 12, 11), VerticalAlignment.BOTTOM, lines);
    }

    private XSSFWorkbook reopen(MdFontSettings fonts, VerticalAlignment alignment, String... lines) throws Exception {
        Path path = temporaryFolder.newFile("saved.xlsx").toPath();
        try (XSSFWorkbook workbook = new XSSFWorkbook(); OutputStream output = Files.newOutputStream(path)) {
            new MarkdownWorkbookRenderer().render(Arrays.asList(lines).iterator(), workbook, fonts,
                    new MdSheetSettings(40, alignment));
            workbook.write(output);
        }
        try (InputStream input = Files.newInputStream(path)) {
            return new XSSFWorkbook(input);
        }
    }

    private static MdFontSettings customFonts() {
        return new MdFontSettings("Arial", 23, 19, 17, 13);
    }

    private static XSSFCell cell(XSSFWorkbook workbook, String address) {
        CellReference reference = new CellReference(address);
        assertNotNull(address, workbook.getSheet("spec").getRow(reference.getRow()));
        XSSFCell cell = workbook.getSheet("spec").getRow(reference.getRow()).getCell(reference.getCol());
        assertNotNull(address, cell);
        return cell;
    }

    private static void assertText(XSSFWorkbook workbook, String address, ExpectedText... parts) {
        XSSFCell cell = cell(workbook, address);
        StringBuilder expected = new StringBuilder();
        for (ExpectedText part : parts) {
            expected.append(part.value);
        }
        assertEquals(address, expected.toString(), cell.getStringCellValue());
        XSSFRichTextString rich = cell.getRichStringCellValue();
        int index = 0;
        for (ExpectedText part : parts) {
            for (int offset = 0; offset < part.value.length(); offset++, index++) {
                // 書式指定がない文字はセルのフォントを使う。IDの値やrun数は検証しない。
                XSSFFont actual = rich.getFontAtIndex(index);
                if (actual == null) {
                    actual = workbook.getFontAt(cell.getCellStyle().getFontIndex());
                }
                String message = address + " character " + index + " (" + rich.getString().charAt(index) + ")";
                assertEquals(message, part.name, actual.getFontName());
                assertEquals(message, part.size, actual.getFontHeightInPoints());
                assertEquals(message, part.bold, actual.getBold());
                assertEquals(message, part.italic, actual.getItalic());
                if (part.code) {
                    assertNotNull(message, actual.getXSSFColor());
                    assertArrayEquals(message, new byte[] { (byte) 180, 0, 0 }, actual.getXSSFColor().getRGB());
                }
            }
        }
    }

    private static void assertGrayFill(XSSFCellStyle style) {
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertNotNull(style.getFillForegroundXSSFColor());
        assertArrayEquals(new byte[] { (byte) 232, (byte) 232, (byte) 232 },
                style.getFillForegroundXSSFColor().getRGB());
    }

    private static void assertAlignment(XSSFCellStyle style, VerticalAlignment vertical) {
        assertEquals(vertical, style.getVerticalAlignment());
        assertEquals(HorizontalAlignment.GENERAL, style.getAlignment());
        assertEquals(false, style.getWrapText());
    }

    private static ExpectedText plain(String value) {
        return text(value, false, false);
    }

    private static ExpectedText text(String value, boolean bold, boolean italic) {
        return font(value, "Meiryo", 11, bold, italic);
    }

    private static ExpectedText font(String value, String name, int size, boolean bold, boolean italic) {
        return new ExpectedText(value, name, size, bold, italic, false);
    }

    private static ExpectedText code(String value, String name, int size, boolean bold) {
        return new ExpectedText(value, name, size, bold, false, true);
    }

    private static final class ExpectedText {
        final String value;
        final String name;
        final int size;
        final boolean bold;
        final boolean italic;
        final boolean code;

        ExpectedText(String value, String name, int size, boolean bold, boolean italic, boolean code) {
            this.value = value;
            this.name = name;
            this.size = size;
            this.bold = bold;
            this.italic = italic;
            this.code = code;
        }
    }
}
