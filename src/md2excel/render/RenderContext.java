package md2excel.render;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import md2excel.excel.MdStyleCatalog;

public final class RenderContext {

    final Sheet sheet;
    final MdStyleCatalog styles;
    final RenderState st;
    final MarkdownFontCache fontCache;

    public RenderContext(XSSFWorkbook workbook, Sheet sheet, MdStyleCatalog styles, SheetColumnLayout columnLayout,
            int startRowIndex) {

        if (workbook == null) {
            throw new IllegalArgumentException("workbook must not be null");
        }

        if (sheet == null) {
            throw new IllegalArgumentException("sheet must not be null");
        }

        if (styles == null) {
            throw new IllegalArgumentException("styles must not be null");
        }

        if (columnLayout == null) {
            throw new IllegalArgumentException("columnLayout must not be null");
        }

        this.sheet = sheet;
        this.styles = styles;

        this.st = new RenderState(columnLayout, startRowIndex);

        this.fontCache = new MarkdownFontCache(workbook);
    }
}