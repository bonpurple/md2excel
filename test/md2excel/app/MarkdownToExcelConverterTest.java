package md2excel.app;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.Stream;

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
import md2excel.config.MdFontSettings;
import md2excel.config.MdSheetSettings;

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

    @Test
    public void nestedQuotedCodeBlockUsesQuoteDepth() throws Exception {

        String markdown = String.join("\n", "> > ```text", "> > code", "> > ```");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            // 2段引用内のコード本文はE列
            assertCellText(sheet, 1, 4, "code");

            // B列は1段目の引用装飾
            assertEquals(BorderStyle.THICK, cell(sheet, 1, 1).getCellStyle().getBorderLeft());

            // C列は2段目の引用装飾
            assertEquals(BorderStyle.THICK, cell(sheet, 1, 2).getCellStyle().getBorderLeft());

            // D列はコードブロック枠
            assertEquals(BorderStyle.THIN, cell(sheet, 1, 3).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THIN, cell(sheet, 1, 3).getCellStyle().getBorderTop());

            assertEquals(BorderStyle.THIN, cell(sheet, 1, 3).getCellStyle().getBorderBottom());

            // 引用装飾後もコード本文セルの上下枠線を維持する。
            assertEquals(BorderStyle.THIN, cell(sheet, 1, 4).getCellStyle().getBorderTop());

            assertEquals(BorderStyle.THIN, cell(sheet, 1, 4).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void normalCodeBlockAfterQuotedTableHasSeparatorBlank() throws Exception {

        String markdown = String.join("\n", "> | h1 | h2 |", "> | --- | --- |", "> | a | b |", "```text", "code",
                "```");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            // 引用テーブルはC列から開始する。
            assertCellText(sheet, 1, 2, "h1");
            assertCellText(sheet, 1, 3, "h2");
            assertCellText(sheet, 2, 2, "a");
            assertCellText(sheet, 2, 3, "b");

            // B列は引用装飾列。
            assertEquals(BorderStyle.THICK, cell(sheet, 1, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THICK, cell(sheet, 2, 1).getCellStyle().getBorderLeft());

            // 引用テーブル直後に自動空行が挿入される。
            Row separatorRow = sheet.getRow(3);
            assertNotNull(separatorRow);

            // 通常コードブロックはB列が枠、C列が本文。
            assertCellText(sheet, 4, 2, "code");

            assertEquals(BorderStyle.THIN, cell(sheet, 4, 1).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void horizontalRuleDoesNotUseOuterMarginColumns() throws Exception {

        try (XSSFWorkbook workbook = convert("---")) {
            Sheet sheet = workbook.getSheet("spec");

            Row row = sheet.getRow(1);

            assertNotNull(row);

            // A列は左余白なので、水平線セルを生成しない。
            assertNull(row.getCell(0));

            // B列は最初の描画列。
            assertEquals(BorderStyle.HAIR, cell(sheet, 1, 1).getCellStyle().getBorderBottom());

            // AM列は最後の描画列。
            assertEquals(BorderStyle.HAIR, cell(sheet, 1, 38).getCellStyle().getBorderBottom());

            // AN列は右余白なので、水平線セルを生成しない。
            assertNull(row.getCell(39));
        }
    }

    @Test
    public void existingOutputFileIsReplacedAfterSuccessfulWrite() throws Exception {

        Path inputPath = temporaryFolder.newFile("replace-input.md").toPath();

        Path outputPath = temporaryFolder.getRoot().toPath().resolve("replace-output.xlsx");

        Files.write(inputPath, "replacement content".getBytes(StandardCharsets.UTF_8));

        // 既存ファイルを用意する。
        Files.write(outputPath, "old content".getBytes(StandardCharsets.UTF_8));

        new MarkdownToExcelConverter().convert(createConfig(inputPath, outputPath));

        try (InputStream input = Files.newInputStream(outputPath); XSSFWorkbook workbook = new XSSFWorkbook(input)) {

            Sheet sheet = workbook.getSheet("spec");

            assertCellText(sheet, 1, 1, "replacement content");
        }

        assertEquals(0L, countTemporaryOutputFiles(outputPath.getParent()));
    }

    @Test
    public void temporaryOutputIsDeletedWhenReplacementFails() throws Exception {

        Path inputPath = temporaryFolder.newFile("failure-input.md").toPath();

        Files.write(inputPath, "content".getBytes(StandardCharsets.UTF_8));

        Path outputPath = temporaryFolder.getRoot().toPath().resolve("failure-output.xlsx");

        // 出力先と同じ名前の空でないディレクトリを作る。
        Files.createDirectory(outputPath);

        Path existingChild = outputPath.resolve("keep.txt");

        Files.write(existingChild, "keep".getBytes(StandardCharsets.UTF_8));

        try {
            new MarkdownToExcelConverter().convert(createConfig(inputPath, outputPath));

            fail("IOException was expected");

        } catch (IOException expected) {
            // expected
        }

        // 元のディレクトリと内容が維持されている。
        assertTrue(Files.isDirectory(outputPath));
        assertTrue(Files.exists(existingChild));

        // 失敗時に一時ファイルが残っていない。
        assertEquals(0L, countTemporaryOutputFiles(outputPath.getParent()));
    }

    @Test
    public void quotedTableAndNormalTableUseSeparateContexts() throws Exception {

        String markdown = String.join("\n", "> | quoted-h1 | quoted-h2 |", "> | --- | --- |",
                "> | quoted-a | quoted-b |", "| normal-h1 | normal-h2 |", "| --- | --- |", "| normal-a | normal-b |");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            /*
             * 引用テーブル: B列は引用装飾、内容はC列から。
             */
            int quotedHeaderRow = findRowWithText(sheet, 2, "quoted-h1");

            int quotedBodyRow = findRowWithText(sheet, 2, "quoted-a");

            /*
             * 通常テーブル: 内容はB列から。
             */
            int normalHeaderRow = findRowWithText(sheet, 1, "normal-h1");

            int normalBodyRow = findRowWithText(sheet, 1, "normal-a");

            assertTrue(quotedHeaderRow < quotedBodyRow);
            assertTrue(quotedBodyRow < normalHeaderRow);
            assertTrue(normalHeaderRow < normalBodyRow);

            assertCellText(sheet, quotedHeaderRow, 3, "quoted-h2");

            assertCellText(sheet, normalHeaderRow, 2, "normal-h2");

            assertEquals(BorderStyle.THIN, cell(sheet, quotedHeaderRow, 2).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.THIN, cell(sheet, normalHeaderRow, 1).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.NONE, cell(sheet, quotedBodyRow, 2).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.NONE, cell(sheet, normalBodyRow, 1).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void tablesAtDifferentQuoteDepthsUseSeparateContexts() throws Exception {

        String markdown = String.join("\n", "> | outer-h1 | outer-h2 |", "> | --- | --- |", "> | outer-a | outer-b |",
                "> > | inner-h1 | inner-h2 |", "> > | --- | --- |", "> > | inner-a | inner-b |");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            // 1段引用のテーブル内容はC列から。
            int outerHeaderRow = findRowWithText(sheet, 2, "outer-h1");

            int outerBodyRow = findRowWithText(sheet, 2, "outer-a");

            // 2段引用のテーブル内容はD列から。
            int innerHeaderRow = findRowWithText(sheet, 3, "inner-h1");

            int innerBodyRow = findRowWithText(sheet, 3, "inner-a");

            assertTrue(outerHeaderRow < outerBodyRow);
            assertTrue(outerBodyRow < innerHeaderRow);
            assertTrue(innerHeaderRow < innerBodyRow);

            assertCellText(sheet, outerHeaderRow, 3, "outer-h2");

            assertCellText(sheet, innerHeaderRow, 4, "inner-h2");

            assertEquals(BorderStyle.THIN, cell(sheet, outerHeaderRow, 2).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.THIN, cell(sheet, innerHeaderRow, 3).getCellStyle().getBorderBottom());
        }
    }

    @Test
    public void nestedQuoteCanReturnToOuterQuote() throws Exception {

        String markdown = String.join("\n", "> > inner paragraph", "> outer paragraph");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            // 2段引用の本文はD列。
            int innerRow = findRowWithText(sheet, 3, "inner paragraph");

            // 1段引用の本文はC列。
            int outerRow = findRowWithText(sheet, 2, "outer paragraph");

            assertTrue(innerRow < outerRow);

            // 内側の行はB列・C列の両方が引用装飾列。
            assertEquals(BorderStyle.THICK, cell(sheet, innerRow, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THICK, cell(sheet, innerRow, 2).getCellStyle().getBorderLeft());

            // 外側へ戻った行ではB列だけが引用装飾列。
            assertEquals(BorderStyle.THICK, cell(sheet, outerRow, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.NONE, cell(sheet, outerRow, 2).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void quotedParagraphAfterQuotedCodeKeepsQuoteContext() throws Exception {

        String markdown = String.join("\n", "> ```text", "> code body", "> ```", "> paragraph after code");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            // 引用コード本文はD列。
            int codeRow = findRowWithText(sheet, 3, "code body");

            // 1段引用の通常本文はC列。
            int paragraphRow = findRowWithText(sheet, 2, "paragraph after code");

            assertTrue(codeRow < paragraphRow);

            // B列は引用装飾。
            assertEquals(BorderStyle.THICK, cell(sheet, codeRow, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THICK, cell(sheet, paragraphRow, 1).getCellStyle().getBorderLeft());

            // C列はコードブロック枠。
            assertEquals(BorderStyle.THIN, cell(sheet, codeRow, 2).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void nestedBulletListCanReturnToRootDepth() throws Exception {

        String markdown = String.join("\n", "- root item", "  - child item", "- root sibling");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            // ルート箇条書きはC列。
            int rootRow = findRowWithText(sheet, 2, "・ root item");

            // 子箇条書きはD列。
            int childRow = findRowWithText(sheet, 3, "・ child item");

            // 浅い階層へ戻るとC列。
            int siblingRow = findRowWithText(sheet, 2, "・ root sibling");

            assertTrue(rootRow < childRow);
            assertTrue(childRow < siblingRow);
        }
    }

    @Test
    public void nestedQuotedCodeBlockAtEofClosesFrame() throws Exception {

        String markdown = String.join("\n", "> > ```text", "> > line1", "> > line2");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            // 2段引用内コード本文はE列。
            int firstRow = findRowWithText(sheet, 4, "line1");

            int lastRow = findRowWithText(sheet, 4, "line2");

            assertTrue(firstRow < lastRow);

            // B列・C列は引用装飾。
            assertEquals(BorderStyle.THICK, cell(sheet, firstRow, 1).getCellStyle().getBorderLeft());

            assertEquals(BorderStyle.THICK, cell(sheet, firstRow, 2).getCellStyle().getBorderLeft());

            // D列はコード枠。
            assertEquals(BorderStyle.THIN, cell(sheet, firstRow, 3).getCellStyle().getBorderTop());

            assertEquals(BorderStyle.THIN, cell(sheet, lastRow, 3).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.THIN, cell(sheet, lastRow, 3).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void normalTableAndQuotedTableUseSeparateContexts() throws Exception {

        String markdown = String.join("\n", "| normal-h1 | normal-h2 |", "| --- | --- |", "| normal-a | normal-b |",
                "> | quoted-h1 | quoted-h2 |", "> | --- | --- |", "> | quoted-a | quoted-b |");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            int normalHeaderRow = findRowWithText(sheet, 1, "normal-h1");

            int normalBodyRow = findRowWithText(sheet, 1, "normal-a");

            int quotedHeaderRow = findRowWithText(sheet, 2, "quoted-h1");

            int quotedBodyRow = findRowWithText(sheet, 2, "quoted-a");

            assertTrue(normalHeaderRow < normalBodyRow);
            assertTrue(normalBodyRow < quotedHeaderRow);
            assertTrue(quotedHeaderRow < quotedBodyRow);

            assertCellText(sheet, normalHeaderRow, 2, "normal-h2");

            assertCellText(sheet, quotedHeaderRow, 3, "quoted-h2");

            assertEquals(BorderStyle.NONE, cell(sheet, normalBodyRow, 1).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.NONE, cell(sheet, quotedBodyRow, 2).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.THICK, cell(sheet, quotedHeaderRow, 1).getCellStyle().getBorderLeft());
        }
    }

    @Test
    public void tableCanReturnFromNestedQuoteToOuterQuote() throws Exception {

        String markdown = String.join("\n", "> > | inner-h1 | inner-h2 |", "> > | --- | --- |",
                "> > | inner-a | inner-b |", "> | outer-h1 | outer-h2 |", "> | --- | --- |", "> | outer-a | outer-b |");

        try (XSSFWorkbook workbook = convert(markdown)) {
            Sheet sheet = workbook.getSheet("spec");

            int innerHeaderRow = findRowWithText(sheet, 3, "inner-h1");

            int innerBodyRow = findRowWithText(sheet, 3, "inner-a");

            int outerHeaderRow = findRowWithText(sheet, 2, "outer-h1");

            int outerBodyRow = findRowWithText(sheet, 2, "outer-a");

            assertTrue(innerHeaderRow < innerBodyRow);
            assertTrue(innerBodyRow < outerHeaderRow);
            assertTrue(outerHeaderRow < outerBodyRow);

            assertCellText(sheet, innerHeaderRow, 4, "inner-h2");

            assertCellText(sheet, outerHeaderRow, 3, "outer-h2");

            assertEquals(BorderStyle.THIN, cell(sheet, innerHeaderRow, 3).getCellStyle().getBorderBottom());

            assertEquals(BorderStyle.THIN, cell(sheet, outerHeaderRow, 2).getCellStyle().getBorderBottom());
        }
    }

    private XSSFWorkbook convert(String markdown) throws Exception {

        Path markdownPath = temporaryFolder.newFile("input.md").toPath();

        Path excelPath = temporaryFolder.getRoot().toPath().resolve("output.xlsx");

        Files.write(markdownPath, markdown.getBytes(StandardCharsets.UTF_8));

        Md2ExcelConfig config = createConfig(markdownPath, excelPath);

        new MarkdownToExcelConverter().convert(config);

        try (InputStream input = Files.newInputStream(excelPath)) {

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

    private static long countTemporaryOutputFiles(Path directory) throws IOException {

        try (Stream<Path> paths = Files.list(directory)) {

            return paths.filter(path -> path.getFileName().toString().startsWith(".md2excel-")).count();
        }
    }

    private static Md2ExcelConfig createConfig(Path inputPath, Path outputPath) {

        return new Md2ExcelConfig(inputPath, outputPath, new MdFontSettings("Meiryo", 16, 14, 12, 11),
                new MdSheetSettings(40, VerticalAlignment.BOTTOM));
    }

    private static int findRowWithText(Sheet sheet, int columnIndex, String expected) {

        for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {

            Row row = sheet.getRow(rowIndex);

            if (row == null) {
                continue;
            }

            Cell cell = row.getCell(columnIndex);

            if (cell == null) {
                continue;
            }

            if (cell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING) {

                continue;
            }

            if (expected.equals(cell.getStringCellValue())) {

                return rowIndex;
            }
        }

        fail("Text was not found: col=" + columnIndex + ", text=" + expected);

        return -1;
    }
}