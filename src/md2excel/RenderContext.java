package md2excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

final class RenderContext {
    final Workbook wb;
    final Sheet sheet;
    final MdStyle styles;
    final RenderState st;

    RenderContext(Workbook wb, Sheet sheet, MdStyle styles, int mergeCols) {
        this.wb = wb;
        this.sheet = sheet;
        this.styles = styles;
        this.st = new RenderState(mergeCols);
    }
}
