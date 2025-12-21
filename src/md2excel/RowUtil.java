package md2excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public final class RowUtil {
    private RowUtil() {
    }

    public enum ReuseKind {
        CODE_LINE,
        TABLE_ROW,
        HORIZONTAL_RULE,
        BLOCK_QUOTE,
        BULLET_ITEM,
        NUMBER_ITEM
    }

    public static Row createRow(Sheet sheet, RenderState st, CellStyle defaultRowStyle) {
        Row row = sheet.createRow(st.rowIndex++);
        ensureRowStyle(row, defaultRowStyle);
        return row;
    }

    private static Row createRowOrReusePreviousMarkdownBlank(Sheet sheet, RenderState st, boolean canReuseBlank,
            CellStyle defaultRowStyle) {

        Row row;
        if (st.lastRowType == RenderState.RowType.BLANK && st.lastBlankFromMarkdown && st.rowIndex > 0
                && canReuseBlank) {
            row = sheet.getRow(st.rowIndex - 1);
            if (row == null)
                row = sheet.createRow(st.rowIndex - 1);
        } else {
            row = sheet.createRow(st.rowIndex++);
        }
        ensureRowStyle(row, defaultRowStyle);
        return row;
    }

    // canReuseBlank判定を RowUtil に寄せた版
    public static Row createRowOrReusePreviousMarkdownBlank(Sheet sheet, RenderState st, ReuseKind kind,
            CellStyle defaultRowStyle) {
        boolean canReuseBlank = canReuseMarkdownBlank(st, kind);
        return createRowOrReusePreviousMarkdownBlank(sheet, st, canReuseBlank, defaultRowStyle);
    }

    // 判定ルールをここに集約（仕様維持）
    private static boolean canReuseMarkdownBlank(RenderState st, ReuseKind kind) {
        if (st.lastBlankAfterTable)
            return false;
        switch (kind) {
        case HORIZONTAL_RULE:
            return true;

        case BLOCK_QUOTE:
            // 既存：見出し直後は詰めない＋引用の連続扱いも詰めない
            return st.lastContentType != RenderState.ContentType.HEADING && !st.lastWasBlockQuote;

        case BULLET_ITEM:
        case NUMBER_ITEM:
            // 既存：見出し直後は詰めない
            return st.lastContentType != RenderState.ContentType.HEADING;

        case CODE_LINE:
        case TABLE_ROW:
            // 既存：直前コンテンツが (BULLET/NUMBER/NORMAL) のときだけ詰める
            return st.lastContentType == RenderState.ContentType.BULLET
                    || st.lastContentType == RenderState.ContentType.NUMBER
                    || st.lastContentType == RenderState.ContentType.NORMAL;

        default:
            return false;
        }
    }

    public static Row reuseLastMarkdownBlankRow(Sheet sheet, RenderState st, CellStyle defaultRowStyle) {
        Row row;
        if (st.lastBlankRowIndex < 0) {
            row = sheet.createRow(st.rowIndex++);
        } else {
            row = sheet.getRow(st.lastBlankRowIndex);
            if (row == null)
                row = sheet.createRow(st.lastBlankRowIndex);
        }
        ensureRowStyle(row, defaultRowStyle);
        return row;
    }

    // 任意 rowNum を get or create（st.rowIndex は触らない）
    public static Row getOrCreateRow(Sheet sheet, int rowNum, CellStyle defaultRowStyle) {
        Row row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        ensureRowStyle(row, defaultRowStyle);
        return row;
    }

    public static void ensureRowStyle(Row row, CellStyle defaultRowStyle) {
        if (row.getRowStyle() == null) {
            row.setRowStyle(defaultRowStyle);
        }
    }

    // =========================
    // ctx版オーバーロード
    // =========================
    public static Row createRow(RenderContext ctx, CellStyle defaultRowStyle) {
        return createRow(ctx.sheet, ctx.st, defaultRowStyle);
    }

    public static Row createRowOrReusePreviousMarkdownBlank(RenderContext ctx, ReuseKind kind,
            CellStyle defaultRowStyle) {
        return createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, kind, defaultRowStyle);
    }
}
