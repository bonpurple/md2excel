package md2excel.render;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import md2excel.excel.MdStyleCatalog;

public final class RenderContext {
    final Workbook wb;
    final Sheet sheet;
    final MdStyleCatalog styles;
    final RenderState st;
    final MarkdownFontCache fontCache;

    public RenderContext(Workbook wb, Sheet sheet, MdStyleCatalog styles, int sheetColumnCount) {

        this(wb, sheet, styles, sheetColumnCount, 0, 0);
    }

    public RenderContext(Workbook wb, Sheet sheet, MdStyleCatalog styles, int sheetColumnCount, int startRowIndex,
            int startColIndex) {

        this.wb = wb;
        this.sheet = sheet;
        this.styles = styles;
        this.st = new RenderState(sheetColumnCount, startRowIndex, startColIndex);

        this.fontCache = new MarkdownFontCache(wb);
    }
}