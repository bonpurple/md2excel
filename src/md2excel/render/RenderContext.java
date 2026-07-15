package md2excel.render;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import md2excel.excel.MdStyle;

public final class RenderContext {
    final Workbook wb;
    final Sheet sheet;
    final MdStyle styles;
    final RenderState st;

    public RenderContext(Workbook wb, Sheet sheet, MdStyle styles, int mergeCols) {
        this(wb, sheet, styles, mergeCols, 0, 0);
    }

    public RenderContext(Workbook wb, Sheet sheet, MdStyle styles, int mergeCols, int startRowIndex,
            int startColIndex) {
        this.wb = wb;
        this.sheet = sheet;
        this.styles = styles;
        this.st = new RenderState(mergeCols, startRowIndex, startColIndex);
    }
}