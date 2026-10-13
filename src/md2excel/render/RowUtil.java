package md2excel.render;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

final class RowUtil {

    private RowUtil() {
    }

    static Row createRow(Sheet sheet, RenderState st, CellStyle defaultRowStyle) {
        Row row = sheet.createRow(st.allocateNextRowIndex());
        ensureRowStyle(row, defaultRowStyle);
        return row;
    }

    /**
     * 水平線用。直前に再利用可能なMarkdown空行があれば、その行を使う。 なければ新しい行を作る。
     */
    static Row createRowOrReusePreviousMarkdownBlankForRule(RenderContext ctx, CellStyle defaultRowStyle) {

        if (!ctx.st.hasPreviousReusableMarkdownBlankRow()) {
            return createRow(ctx.sheet, ctx.st, defaultRowStyle);
        }

        return getOrCreateRow(ctx.sheet, ctx.st.getPreviousRowIndex(), defaultRowStyle);
    }

    /**
     * 段落用の空行再利用。水平線用の規則とは別に維持する。
     */
    static Row reuseLastMarkdownBlankRow(Sheet sheet, RenderState st, CellStyle defaultRowStyle) {

        Row row;

        int blankRowIndex = st.getLastBlankRowIndex();

        if (blankRowIndex < 0) {
            row = sheet.createRow(st.allocateNextRowIndex());
        } else {
            row = sheet.getRow(blankRowIndex);

            if (row == null) {
                row = sheet.createRow(blankRowIndex);
            }
        }

        ensureRowStyle(row, defaultRowStyle);
        return row;
    }

    private static Row getOrCreateRow(Sheet sheet, int rowNum, CellStyle defaultRowStyle) {

        Row row = sheet.getRow(rowNum);

        if (row == null) {
            row = sheet.createRow(rowNum);
        }

        ensureRowStyle(row, defaultRowStyle);
        return row;
    }

    private static void ensureRowStyle(Row row, CellStyle defaultRowStyle) {
        if (row.getRowStyle() == null) {
            row.setRowStyle(defaultRowStyle);
        }
    }
}