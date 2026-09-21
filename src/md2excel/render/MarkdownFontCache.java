package md2excel.render;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import md2excel.excel.MdStyleDefaults;

public final class MarkdownFontCache {

    private final Workbook workbook;

    private final Map<Short, InlineFonts> inlineFontsByBaseFontIndex = new HashMap<Short, InlineFonts>();

    private final Map<Short, CodeBlockFonts> codeBlockFontsByStyleFontIndex = new HashMap<Short, CodeBlockFonts>();

    public MarkdownFontCache(Workbook workbook) {
        if (workbook == null) {
            throw new IllegalArgumentException("workbook must not be null");
        }

        this.workbook = workbook;
    }

    InlineFonts getInlineFonts(CellStyle baseStyle) {
        short key = (short) baseStyle.getFontIndex();

        InlineFonts cached = inlineFontsByBaseFontIndex.get(Short.valueOf(key));

        if (cached != null) {
            return cached;
        }

        Font baseFont = workbook.getFontAt(baseStyle.getFontIndex());
        boolean baseBold = baseFont.getBold();

        Font boldFont = createFont(baseFont.getFontName(), baseFont.getFontHeightInPoints(), true, false);

        Font italicFont = createFont(baseFont.getFontName(), baseFont.getFontHeightInPoints(), baseBold, true);

        Font boldItalicFont = createFont(baseFont.getFontName(), baseFont.getFontHeightInPoints(), true, true);

        XSSFColor inlineCodeColor = new XSSFColor(MdStyleDefaults.INLINE_CODE_TEXT, null);

        XSSFFont codeAscii = createXssfFont(MdStyleDefaults.CODE_ASCII_FONT_NAME, baseFont.getFontHeightInPoints(),
                false, inlineCodeColor);

        XSSFFont codeCjk = createXssfFont(MdStyleDefaults.CODE_CJK_FONT_NAME, baseFont.getFontHeightInPoints(), false,
                inlineCodeColor);

        XSSFFont codeAsciiBold = createXssfFont(MdStyleDefaults.CODE_ASCII_FONT_NAME, baseFont.getFontHeightInPoints(),
                true, inlineCodeColor);

        XSSFFont codeCjkBold = createXssfFont(MdStyleDefaults.CODE_CJK_FONT_NAME, baseFont.getFontHeightInPoints(),
                true, inlineCodeColor);

        InlineFonts fonts = new InlineFonts(baseFont, boldFont, italicFont, boldItalicFont, codeAscii, codeCjk,
                codeAsciiBold, codeCjkBold, baseBold);

        inlineFontsByBaseFontIndex.put(Short.valueOf(key), fonts);
        return fonts;
    }

    CodeBlockFonts getCodeBlockFonts(CellStyle codeBlockStyle) {
        short key = (short) codeBlockStyle.getFontIndex();

        CodeBlockFonts cached = codeBlockFontsByStyleFontIndex.get(Short.valueOf(key));

        if (cached != null) {
            return cached;
        }

        Font baseFont = workbook.getFontAt(codeBlockStyle.getFontIndex());

        Font asciiFont = createFont(MdStyleDefaults.CODE_ASCII_FONT_NAME, baseFont.getFontHeightInPoints(), false,
                false);

        Font cjkFont = createFont(MdStyleDefaults.CODE_CJK_FONT_NAME, baseFont.getFontHeightInPoints(), false, false);

        CodeBlockFonts fonts = new CodeBlockFonts(asciiFont, cjkFont);

        codeBlockFontsByStyleFontIndex.put(Short.valueOf(key), fonts);

        return fonts;
    }

    private Font createFont(String fontName, short fontHeight, boolean bold, boolean italic) {

        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeightInPoints(fontHeight);
        font.setBold(bold);
        font.setItalic(italic);

        return font;
    }

    private XSSFFont createXssfFont(String fontName, short fontHeight, boolean bold, XSSFColor color) {

        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeightInPoints(fontHeight);
        font.setBold(bold);
        font.setColor(color);

        return font;
    }

    static final class InlineFonts {
        final Font baseFont;
        final Font boldFont;
        final Font italicFont;
        final Font boldItalicFont;

        final XSSFFont codeAscii;
        final XSSFFont codeCjk;
        final XSSFFont codeAsciiBold;
        final XSSFFont codeCjkBold;

        final boolean baseBold;

        InlineFonts(Font baseFont, Font boldFont, Font italicFont, Font boldItalicFont, XSSFFont codeAscii,
                XSSFFont codeCjk, XSSFFont codeAsciiBold, XSSFFont codeCjkBold, boolean baseBold) {

            this.baseFont = baseFont;
            this.boldFont = boldFont;
            this.italicFont = italicFont;
            this.boldItalicFont = boldItalicFont;

            this.codeAscii = codeAscii;
            this.codeCjk = codeCjk;
            this.codeAsciiBold = codeAsciiBold;
            this.codeCjkBold = codeCjkBold;

            this.baseBold = baseBold;
        }
    }

    static final class CodeBlockFonts {
        final Font ascii;
        final Font cjk;

        CodeBlockFonts(Font ascii, Font cjk) {
            this.ascii = ascii;
            this.cjk = cjk;
        }
    }
}