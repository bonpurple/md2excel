package md2excel.render;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import md2excel.excel.MdStyleDefaults;

public final class MarkdownFontCache {

    private final XSSFWorkbook workbook;

    private final Map<Short, InlineFonts> inlineFontsByBaseFontIndex = new HashMap<Short, InlineFonts>();

    private final Map<Short, CodeBlockFonts> codeBlockFontsByStyleFontIndex = new HashMap<Short, CodeBlockFonts>();

    public MarkdownFontCache(XSSFWorkbook workbook) {

        if (workbook == null) {
            throw new IllegalArgumentException("workbook must not be null");
        }

        this.workbook = workbook;
    }

    InlineFonts getInlineFonts(CellStyle baseStyle) {

        if (baseStyle == null) {
            throw new IllegalArgumentException("baseStyle must not be null");
        }

        short key = (short) baseStyle.getFontIndex();

        InlineFonts cached = inlineFontsByBaseFontIndex.get(Short.valueOf(key));

        if (cached != null) {
            return cached;
        }

        InlineFonts fonts = createInlineFonts(baseStyle);

        inlineFontsByBaseFontIndex.put(Short.valueOf(key), fonts);

        return fonts;
    }

    private InlineFonts createInlineFonts(CellStyle baseStyle) {

        XSSFFont baseFont = workbook.getFontAt(baseStyle.getFontIndex());

        boolean baseBold = baseFont.getBold();

        XSSFFont boldFont = createFont(baseFont.getFontName(), baseFont.getFontHeightInPoints(), true, false);

        XSSFFont italicFont = createFont(baseFont.getFontName(), baseFont.getFontHeightInPoints(), baseBold, true);

        XSSFFont boldItalicFont = createFont(baseFont.getFontName(), baseFont.getFontHeightInPoints(), true, true);

        XSSFColor inlineCodeColor = new XSSFColor(MdStyleDefaults.INLINE_CODE_TEXT, null);

        XSSFFont codeAscii = createColoredFont(MdStyleDefaults.CODE_ASCII_FONT_NAME, baseFont.getFontHeightInPoints(),
                false, inlineCodeColor);

        XSSFFont codeCjk = createColoredFont(MdStyleDefaults.CODE_CJK_FONT_NAME, baseFont.getFontHeightInPoints(),
                false, inlineCodeColor);

        XSSFFont codeAsciiBold = createColoredFont(MdStyleDefaults.CODE_ASCII_FONT_NAME,
                baseFont.getFontHeightInPoints(), true, inlineCodeColor);

        XSSFFont codeCjkBold = createColoredFont(MdStyleDefaults.CODE_CJK_FONT_NAME, baseFont.getFontHeightInPoints(),
                true, inlineCodeColor);

        return new InlineFonts(baseFont, boldFont, italicFont, boldItalicFont, codeAscii, codeCjk, codeAsciiBold,
                codeCjkBold, baseBold);
    }

    CodeBlockFonts getCodeBlockFonts(CellStyle codeBlockStyle) {

        if (codeBlockStyle == null) {
            throw new IllegalArgumentException("codeBlockStyle must not be null");
        }

        short key = (short) codeBlockStyle.getFontIndex();

        CodeBlockFonts cached = codeBlockFontsByStyleFontIndex.get(Short.valueOf(key));

        if (cached != null) {
            return cached;
        }

        CodeBlockFonts fonts = createCodeBlockFonts(codeBlockStyle);

        codeBlockFontsByStyleFontIndex.put(Short.valueOf(key), fonts);

        return fonts;
    }

    private CodeBlockFonts createCodeBlockFonts(CellStyle codeBlockStyle) {

        XSSFFont baseFont = workbook.getFontAt(codeBlockStyle.getFontIndex());

        XSSFFont asciiFont = createFont(MdStyleDefaults.CODE_ASCII_FONT_NAME, baseFont.getFontHeightInPoints(), false,
                false);

        XSSFFont cjkFont = createFont(MdStyleDefaults.CODE_CJK_FONT_NAME, baseFont.getFontHeightInPoints(), false,
                false);

        return new CodeBlockFonts(asciiFont, cjkFont);
    }

    private XSSFFont createFont(String fontName, short fontHeight, boolean bold, boolean italic) {

        XSSFFont font = workbook.createFont();

        font.setFontName(fontName);
        font.setFontHeightInPoints(fontHeight);
        font.setBold(bold);
        font.setItalic(italic);

        return font;
    }

    private XSSFFont createColoredFont(String fontName, short fontHeight, boolean bold, XSSFColor color) {

        XSSFFont font = createFont(fontName, fontHeight, bold, false);

        font.setColor(color);

        return font;
    }

    static final class InlineFonts {

        final XSSFFont baseFont;
        final XSSFFont boldFont;
        final XSSFFont italicFont;
        final XSSFFont boldItalicFont;

        final XSSFFont codeAscii;
        final XSSFFont codeCjk;
        final XSSFFont codeAsciiBold;
        final XSSFFont codeCjkBold;

        final boolean baseBold;

        InlineFonts(XSSFFont baseFont, XSSFFont boldFont, XSSFFont italicFont, XSSFFont boldItalicFont,
                XSSFFont codeAscii, XSSFFont codeCjk, XSSFFont codeAsciiBold, XSSFFont codeCjkBold, boolean baseBold) {

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

        final XSSFFont ascii;
        final XSSFFont cjk;

        CodeBlockFonts(XSSFFont ascii, XSSFFont cjk) {

            this.ascii = ascii;
            this.cjk = cjk;
        }
    }
}
