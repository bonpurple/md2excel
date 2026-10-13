package md2excel.render;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import md2excel.excel.ExcelCellUtil;
import md2excel.excel.MdStyleCatalog;

final class BlockQuoteUtil {
    private BlockQuoteUtil() {
    }

    static void closeBlockQuoteIfOpen(Sheet sheet, MdStyleCatalog styles, RenderState st) {

        if (!st.hasRenderableBlockQuote()) {
            st.clearBlockQuoteTracking();
            return;
        }

        applyBlockQuoteStyle(sheet, styles, st, st.getBlockQuoteFirstRow(), st.getBlockQuoteLastRow(), st.getBlockQuoteStartCol(), st.getRenderLastColIndex());

        st.clearBlockQuoteTracking();
    }

    private static void applyBlockQuoteStyle(Sheet sheet, MdStyleCatalog styles, RenderState st, int firstRow, int lastRow, int startCol, int lastColIndex) {

        int fillEndCol = Math.max(startCol, lastColIndex);

        for (int r = firstRow; r <= lastRow; r++) {
            Row rowObj = sheet.getRow(r);

            if (rowObj == null) {
                continue;
            }

            QuoteRowInfo quoteRowInfo = st.getBlockQuoteRowInfo(r);

            QuoteRowKind quoteRowKind = quoteRowInfo == null ? QuoteRowKind.NORMAL : quoteRowInfo.getKind();

            int quoteDepth = quoteRowInfo == null ? 1 : quoteRowInfo.getDepth();

            for (int c = startCol; c <= fillEndCol; c++) {
                Cell cell = ExcelCellUtil.getOrCreateCell(rowObj, c);

                applyBlockQuoteCellStyle(cell, styles, quoteRowInfo, quoteRowKind, quoteDepth, startCol, c);
            }
        }
    }

    private static void applyBlockQuoteCellStyle(Cell cell, MdStyleCatalog styles, QuoteRowInfo quoteRowInfo, QuoteRowKind quoteRowKind,
        int quoteDepth, int startCol, int col) {

        // startColは最も左側の引用装飾列。
        // 引用深度分の列を引用罫線列として扱う。
        boolean isQuoteDecorCol = col >= startCol && col < startCol + quoteDepth;

        // ----------------------------------------
        // code
        // ----------------------------------------
        if (quoteRowKind == QuoteRowKind.CODE) {
            if (isQuoteDecorCol) {
                cell.setCellStyle(styles.blockQuoteLeftStyle);
            }

            // 引用装飾列以外はcodeBlockFrameStyleを維持する。
            return;
        }

        // ----------------------------------------
        // horizontal rule
        // ----------------------------------------
        if (quoteRowKind == QuoteRowKind.HORIZONTAL_RULE) {

            if (isQuoteDecorCol) {
                cell.setCellStyle(styles.blockQuoteBlankLeftStyle);
            } else {
                cell.setCellStyle(styles.blockQuoteHorizontalRuleBodyStyle);
            }

            return;
        }

        // ----------------------------------------
        // table
        // ----------------------------------------
        if (quoteRowKind == QuoteRowKind.TABLE) {
            if (isQuoteDecorCol) {
                cell.setCellStyle(styles.blockQuoteLeftStyle);

            } else if (quoteRowInfo != null && quoteRowInfo.isTableContentColumn(col)) {

                cell.setCellStyle(resolveQuoteTableStyle(quoteRowInfo.getTableRowStyleRole(), styles));

            } else {
                // テーブル範囲より右側は引用背景だけを適用する。
                cell.setCellStyle(styles.blockQuoteBodyStyle);
            }

            return;
        }

        // ----------------------------------------
        // blank
        // ----------------------------------------
        if (quoteRowKind == QuoteRowKind.BLANK) {
            cell.setCellStyle(isQuoteDecorCol ? styles.blockQuoteBlankLeftStyle : styles.blockQuoteBlankBodyStyle);

            return;
        }

        // ----------------------------------------
        // normal / heading / list
        // ----------------------------------------
        if (isQuoteDecorCol) {
            cell.setCellStyle(styles.blockQuoteLeftStyle);
        } else {
            cell.setCellStyle(resolveBlockQuoteContentStyle(quoteRowInfo, col, styles));
        }
    }

    private static CellStyle resolveBlockQuoteContentStyle(QuoteRowInfo quoteRowInfo, int col, MdStyleCatalog styles) {

        if (quoteRowInfo == null || col != quoteRowInfo.getContentCol()) {
            return styles.blockQuoteBodyStyle;
        }

        switch (quoteRowInfo.getKind()) {
        case HEADING_1:
            return styles.blockQuoteHeading1Style;

        case HEADING_2:
            return styles.blockQuoteHeading2Style;

        case HEADING_3:
            return styles.blockQuoteHeading3Style;

        case HEADING_4:
            return styles.blockQuoteHeading4Style;

        default:
            return styles.blockQuoteBodyStyle;
        }
    }

    private static CellStyle resolveQuoteTableStyle(TableRowStyleRole role, MdStyleCatalog styles) {

        if (role == null) {
            return styles.blockQuoteBodyStyle;
        }

        switch (role) {
        case HEADER:
            return styles.tableHeaderQuoteStyle;

        case BODY_WITH_BOTTOM_BORDER:
            return styles.tableBodyQuoteStyle;

        case BODY_WITHOUT_BOTTOM_BORDER:
            return styles.tableBodyLastRowQuoteStyle;

        default:
            return styles.blockQuoteBodyStyle;
        }
    }
}