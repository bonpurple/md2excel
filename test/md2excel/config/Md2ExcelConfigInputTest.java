package md2excel.config;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.nio.file.InvalidPathException;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.Test;

public class Md2ExcelConfigInputTest {

    @Test
    public void createsConfigFromInputValues() {
        Md2ExcelConfig config = create(" work/document.md ", " 游ゴシック ", VerticalAlignment.BOTTOM, 16, 14,
                12, 11, 40);

        assertEquals(Paths.get("work", "document.md").toAbsolutePath(), config.getInputPath());
        assertEquals(Paths.get("work", "document.xlsx").toAbsolutePath(), config.getOutputPath());
        assertEquals(new MdFontSettings("游ゴシック", 16, 14, 12, 11), config.getFontSettings());
        assertEquals(new MdSheetSettings(40, VerticalAlignment.BOTTOM), config.getSheetSettings());
    }

    @Test
    public void rejectsEmptyInputPath() {
        assertError("Markdownファイルを選択してください。", "  ", "游ゴシック", VerticalAlignment.BOTTOM,
                16, 14, 12, 11, 40);
    }

    @Test
    public void rejectsInvalidInputPath() {
        try {
            create("bad\0path.md", "游ゴシック", VerticalAlignment.BOTTOM, 16, 14, 12, 11, 40);
            fail("InvalidPathException expected");
        } catch (InvalidPathException expected) {
            // ダイアログでパスの入力エラーとして表示する。
        }
    }

    @Test
    public void rejectsMissingFont() {
        assertError("フォントを選択してください。", "document.md", "  ", VerticalAlignment.BOTTOM,
                16, 14, 12, 11, 40);
    }

    @Test
    public void rejectsMissingAlignment() {
        assertError("セルの縦位置を選択してください。", "document.md", "游ゴシック", null,
                16, 14, 12, 11, 40);
    }

    @Test
    public void rejectsOutOfRangeFontSize() {
        assertError("h1Size must be between 5 and 72: 4", "document.md", "游ゴシック",
                VerticalAlignment.BOTTOM, 4, 14, 12, 11, 40);
    }

    @Test
    public void rejectsOutOfRangeColumnCount() {
        assertError("totalColumnCount must be between 3 and 256: 2", "document.md", "游ゴシック",
                VerticalAlignment.BOTTOM, 16, 14, 12, 11, 2);
    }

    private static Md2ExcelConfig create(String inputText, String fontName, VerticalAlignment alignment, int h1Size,
            int h2Size, int h3Size, int normalSize, int totalColumnCount) {
        return Md2ExcelConfigInput.create(inputText, fontName, alignment, h1Size, h2Size, h3Size, normalSize,
                totalColumnCount);
    }

    private static void assertError(String expectedMessage, String inputText, String fontName,
            VerticalAlignment alignment, int h1Size, int h2Size, int h3Size, int normalSize, int totalColumnCount) {
        try {
            create(inputText, fontName, alignment, h1Size, h2Size, h3Size, normalSize, totalColumnCount);
            fail("IllegalArgumentException expected");
        } catch (IllegalArgumentException expected) {
            assertEquals(expectedMessage, expected.getMessage());
        }
    }
}
