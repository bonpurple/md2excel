package md2excel.render;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * インライン解析とセル書き込みへの既存の入口。
 */
final class MarkdownInline {

    private MarkdownInline() {
    }

    static void setCodeBlockRichTextCell(MarkdownFontCache fontCache, Cell cell, String codeText, CellStyle codeBlockStyle) {

        MarkdownRichTextWriter.setCodeBlockRichTextCell(fontCache, cell, codeText, codeBlockStyle);
    }
}