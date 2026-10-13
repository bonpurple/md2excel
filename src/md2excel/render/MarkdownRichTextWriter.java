package md2excel.render;

import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import md2excel.markdown.MdTextUtil;

/**
 * 解析済みのインライン要素をApache POIのセルへ書き込む。
 */
final class MarkdownRichTextWriter {

    private MarkdownRichTextWriter() {
    }

    static void setResolvedSegmentsCell(MarkdownFontCache fontCache, Cell cell, List<MdSegment> segments, CellStyle baseStyle) {

        if (segments == null) {
            segments = Collections.<MdSegment> emptyList();
        }

        MarkdownFontCache.InlineFonts fonts = fontCache.getInlineFonts(baseStyle);

        XSSFRichTextString rich = new XSSFRichTextString("");
        appendSegmentsToRichText(rich, 0, segments, fonts);

        cell.setCellStyle(baseStyle);
        cell.setCellValue(rich);
    }

    static void setCodeBlockRichTextCell(MarkdownFontCache fontCache, Cell cell, String codeText, CellStyle codeBlockStyle) {

        MarkdownFontCache.CodeBlockFonts fonts = fontCache.getCodeBlockFonts(codeBlockStyle);

        XSSFRichTextString rich = new XSSFRichTextString(codeText);
        applyCodeFonts(rich, codeText, 0, fonts.ascii, fonts.cjk);

        cell.setCellStyle(codeBlockStyle);
        cell.setCellValue(rich);
    }

    private static int appendSegmentsToRichText(XSSFRichTextString rich, int startPos, List<MdSegment> segments, MarkdownFontCache.InlineFonts fonts) {

        int pos = startPos;
        if (segments == null || segments.isEmpty()) {
            return pos;
        }

        for (MdSegment seg : segments) {
            String text = seg.text;
            if (text == null || text.isEmpty()) {
                continue;
            }

            int segStart = pos;
            rich.append(text);
            int segEnd = segStart + text.length();

            if (seg.inCode) {
                boolean wantBoldCode = seg.inBold || fonts.baseBold;
                applyCodeFonts(rich, text, segStart, wantBoldCode ? fonts.codeAsciiBold : fonts.codeAscii, wantBoldCode ? fonts.codeCjkBold : fonts.codeCjk);
            } else if (seg.inBold && seg.inItalic) {
                rich.applyFont(segStart, segEnd, fonts.boldItalicFont);
            } else if (seg.inBold) {
                rich.applyFont(segStart, segEnd, fonts.boldFont);
            } else if (seg.inItalic) {
                rich.applyFont(segStart, segEnd, fonts.italicFont);
            } else {
                rich.applyFont(segStart, segEnd, fonts.baseFont);
            }

            pos = segEnd;
        }

        return pos;
    }

    private static void applyCodeFonts(XSSFRichTextString rich, String text, int startPos, XSSFFont asciiFont, XSSFFont cjkFont) {

        int i = 0;
        while (i < text.length()) {
            int runStart = i;
            boolean ascii = MdTextUtil.isAsciiLike(text.charAt(i));
            i++;

            while (i < text.length() && MdTextUtil.isAsciiLike(text.charAt(i)) == ascii) {
                i++;
            }

            rich.applyFont(startPos + runStart, startPos + i, ascii ? asciiFont : cjkFont);
        }
    }
}