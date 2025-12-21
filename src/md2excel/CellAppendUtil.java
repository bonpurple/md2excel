package md2excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public final class CellAppendUtil {
    private CellAppendUtil() {
    }

    public static void appendMarkdownWithSpace(Sheet sheet, Workbook wb, MdStyle styles, int rowNum, int colNum,
            String markdownText, CellStyle baseStyle) {
        appendMarkdown(sheet, wb, styles, rowNum, colNum, markdownText, baseStyle, true);
    }

    public static void appendMarkdown(Sheet sheet, Workbook wb, MdStyle styles, int rowNum, int colNum,
            String markdownText, CellStyle baseStyle, boolean withLeadingSpace) {

        if (markdownText == null || markdownText.isEmpty()) {
            return;
        }

        // 行は RowUtil に統一
        Row row = RowUtil.getOrCreateRow(sheet, rowNum, styles.normalStyle);

        Cell cell = row.getCell(colNum);
        if (cell == null) {
            cell = row.createCell(colNum);
            // append 前提なので、RichText を空で初期化しておく（安全）
            MarkdownInline.setMarkdownRichTextCell(wb, cell, "", baseStyle);
        }

        MarkdownInline.appendMarkdownToCell(wb, cell, markdownText, baseStyle, withLeadingSpace);
    }

    // =========================
    // ctx版オーバーロード
    // =========================
    public static void appendMarkdownWithSpace(RenderContext ctx, int rowNum, int colNum, String markdownText,
            CellStyle baseStyle) {
        // ここから ctx版へ委譲
        appendMarkdown(ctx, rowNum, colNum, markdownText, baseStyle, true);
    }

    public static void appendMarkdown(RenderContext ctx, int rowNum, int colNum, String markdownText,
            CellStyle baseStyle, boolean withLeadingSpace) {
        // 実体は既存の引数だらけ版へ委譲（挙動は同じ）
        appendMarkdown(ctx.sheet, ctx.wb, ctx.styles, rowNum, colNum, markdownText, baseStyle, withLeadingSpace);
    }
}
