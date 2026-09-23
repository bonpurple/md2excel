package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertSame;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class MarkdownFontCacheTest {

    @Test
    public void reusesInlineFontsForSameBaseStyle() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            XSSFFont baseFont = workbook.createFont();

            baseFont.setFontName("Meiryo");
            baseFont.setFontHeightInPoints((short) 11);

            XSSFCellStyle baseStyle = workbook.createCellStyle();

            baseStyle.setFont(baseFont);

            MarkdownFontCache cache = new MarkdownFontCache(workbook);

            int fontCount = workbook.getNumberOfFonts();

            MarkdownFontCache.InlineFonts first = cache.getInlineFonts(baseStyle);

            assertEquals(fontCount + 7, workbook.getNumberOfFonts());

            MarkdownFontCache.InlineFonts second = cache.getInlineFonts(baseStyle);

            assertSame(first, second);
            assertEquals(fontCount + 7, workbook.getNumberOfFonts());

            assertNotNull(first.baseFont);
            assertNotNull(first.boldFont);
            assertNotNull(first.italicFont);
            assertNotNull(first.boldItalicFont);
            assertNotNull(first.codeAscii);
            assertNotNull(first.codeCjk);
            assertNotNull(first.codeAsciiBold);
            assertNotNull(first.codeCjkBold);
        }
    }

    @Test
    public void reusesCodeBlockFontsForSameStyle() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            XSSFFont baseFont = workbook.createFont();

            baseFont.setFontName("Meiryo");
            baseFont.setFontHeightInPoints((short) 10);

            XSSFCellStyle codeStyle = workbook.createCellStyle();

            codeStyle.setFont(baseFont);

            MarkdownFontCache cache = new MarkdownFontCache(workbook);

            int fontCount = workbook.getNumberOfFonts();

            MarkdownFontCache.CodeBlockFonts first = cache.getCodeBlockFonts(codeStyle);

            assertEquals(fontCount + 2, workbook.getNumberOfFonts());

            MarkdownFontCache.CodeBlockFonts second = cache.getCodeBlockFonts(codeStyle);

            assertSame(first, second);
            assertEquals(fontCount + 2, workbook.getNumberOfFonts());

            assertNotNull(first.ascii);
            assertNotNull(first.cjk);
        }
    }
}
