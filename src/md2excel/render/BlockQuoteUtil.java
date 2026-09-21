package md2excel.render;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import md2excel.excel.ExcelCellUtil;
import md2excel.excel.MdStyleCatalog;

public final class BlockQuoteUtil {
    private BlockQuoteUtil() {
    }

    public static void closeBlockQuoteIfOpen(Sheet sheet, MdStyleCatalog styles, RenderState st) {
        if (!st.inBlockQuote) {
            st.clearBlockQuoteRows();
            return;
        }
        if (st.blockQuoteFirstRow < 0 || st.blockQuoteLastRow < 0) {
            st.clearBlockQuoteRows();
            return;
        }

        applyBlockQuoteStyle(sheet, styles, st, st.blockQuoteFirstRow, st.blockQuoteLastRow, st.blockQuoteCol,
                st.renderLastColIndex);

        st.inBlockQuote = false;
        st.blockQuoteFirstRow = -1;
        st.blockQuoteLastRow = -1;
        st.blockQuoteCellRow = -1;
        st.blockQuoteCellCol = -1;

        st.clearBlockQuoteRows();
    }

    private static void applyBlockQuoteStyle(Sheet sheet, MdStyleCatalog styles, RenderState st, int firstRow,
            int lastRow, int startCol, int lastColIndex) {

        int fillEndCol = Math.max(startCol, lastColIndex);

        for (int r = firstRow; r <= lastRow; r++) {
            Row rowObj = sheet.getRow(r);

            if (rowObj == null) {
                continue;
            }

            RenderState.QuoteRowInfo quoteRowInfo = st.getBlockQuoteRowInfo(r);

            RenderState.QuoteRowKind quoteRowKind = quoteRowInfo == null ? RenderState.QuoteRowKind.NORMAL
                    : quoteRowInfo.kind;

            int quoteDepth = quoteRowInfo == null ? 1 : quoteRowInfo.depth;

            for (int c = startCol; c <= fillEndCol; c++) {
                Cell cell = ExcelCellUtil.getOrCreateCell(rowObj, c);

                // startColは最も左側の引用装飾列。
                // 引用深度分の列を引用罫線列として扱う。
                boolean isQuoteDecorCol = c >= startCol && c < startCol + quoteDepth;

                // ----------------------------------------
                // code
                // ----------------------------------------
                if (quoteRowKind == RenderState.QuoteRowKind.CODE) {
                    if (isQuoteDecorCol) {
                        cell.setCellStyle(styles.blockQuoteLeftStyle);
                    }

                    // 引用装飾列以外はcodeBlockFrameStyleを維持する。
                    continue;
                }

                // ----------------------------------------
                // horizontal rule
                // ----------------------------------------
                if (quoteRowKind == RenderState.QuoteRowKind.HORIZONTAL_RULE) {

                    if (isQuoteDecorCol) {
                        cell.setCellStyle(styles.blockQuoteBlankLeftStyle);
                    } else {
                        cell.setCellStyle(styles.blockQuoteHorizontalRuleBodyStyle);
                    }

                    continue;
                }

                // ----------------------------------------
                // table
                // ----------------------------------------
                if (quoteRowKind == RenderState.QuoteRowKind.TABLE) {
                    if (isQuoteDecorCol) {
                        cell.setCellStyle(styles.blockQuoteLeftStyle);

                    } else if (quoteRowInfo != null && quoteRowInfo.isTableContentColumn(c)) {

                        cell.setCellStyle(resolveQuoteTableStyle(quoteRowInfo.tableRowStyleRole, styles));

                    } else {
                        // テーブル範囲より右側は引用背景だけを適用する。
                        cell.setCellStyle(styles.blockQuoteBodyStyle);
                    }

                    continue;
                }

                // ----------------------------------------
                // blank
                // ----------------------------------------
                if (quoteRowKind == RenderState.QuoteRowKind.BLANK) {
                    cell.setCellStyle(
                            isQuoteDecorCol ? styles.blockQuoteBlankLeftStyle : styles.blockQuoteBlankBodyStyle);

                    continue;
                }

                // ----------------------------------------
                // normal / heading / list
                // ----------------------------------------
                if (isQuoteDecorCol) {
                    cell.setCellStyle(styles.blockQuoteLeftStyle);
                } else {
                    cell.setCellStyle(resolveBlockQuoteContentStyle(quoteRowInfo, c, styles));
                }
            }
        }
    }

    private static CellStyle resolveBlockQuoteContentStyle(RenderState.QuoteRowInfo quoteRowInfo, int col,
            MdStyleCatalog styles) {

        if (quoteRowInfo == null || col != quoteRowInfo.contentCol) {
            return styles.blockQuoteBodyStyle;
        }

        switch (quoteRowInfo.kind) {
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

    private static CellStyle resolveQuoteTableStyle(MarkdownTable.TableRowStyleRole role, MdStyleCatalog styles) {

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