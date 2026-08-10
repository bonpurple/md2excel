package md2excel.render;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import md2excel.excel.MdStyle;

public final class BlockQuoteUtil {
    private BlockQuoteUtil() {
    }

    public static void closeBlockQuoteIfOpen(Sheet sheet, MdStyle styles, RenderState st) {
        if (!st.inBlockQuote) {
            st.clearBlockQuoteRows();
            return;
        }
        if (st.blockQuoteFirstRow < 0 || st.blockQuoteLastRow < 0) {
            st.clearBlockQuoteRows();
            return;
        }

        applyBlockQuoteStyle(sheet, styles, st, st.blockQuoteFirstRow, st.blockQuoteLastRow, st.blockQuoteCol,
                st.lastColIndex);

        st.inBlockQuote = false;
        st.blockQuoteFirstRow = -1;
        st.blockQuoteLastRow = -1;
        st.blockQuoteCellRow = -1;
        st.blockQuoteCellCol = -1;

        st.clearBlockQuoteRows();
    }

    private static void applyBlockQuoteStyle(Sheet sheet, MdStyle styles, RenderState st, int firstRow, int lastRow,
            int startCol, int lastColIndex) {

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
                Cell cell = rowObj.getCell(c);

                if (cell == null) {
                    cell = rowObj.createCell(c);
                    cell.setBlank();
                }

                // ----------------------------------------
                // code
                // ----------------------------------------
                if (quoteRowKind == RenderState.QuoteRowKind.CODE) {

                    if (c == startCol) {
                        cell.setCellStyle(styles.blockQuoteLeftStyle);
                    }

                    // codeBlockFrameStyle は維持する。
                    continue;
                }

                // ----------------------------------------
                // horizontal rule
                // ----------------------------------------
                if (quoteRowKind == RenderState.QuoteRowKind.HORIZONTAL_RULE) {

                    if (c == startCol) {
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

                    if (c == startCol) {
                        cell.setCellStyle(styles.blockQuoteLeftStyle);

                    } else {
                        CellStyle currentStyle = cell.getCellStyle();

                        if (currentStyle.getIndex() == styles.tableHeaderStyle.getIndex()) {

                            cell.setCellStyle(styles.tableHeaderQuoteStyle);

                        } else if (currentStyle.getIndex() == styles.tableBodyLastRowStyle.getIndex()) {

                            cell.setCellStyle(styles.tableBodyLastRowQuoteStyle);

                        } else if (currentStyle.getIndex() == styles.tableBodyStyle.getIndex()) {

                            cell.setCellStyle(styles.tableBodyQuoteStyle);

                        } else {
                            cell.setCellStyle(styles.blockQuoteBodyStyle);
                        }
                    }

                    continue;
                }

                // ----------------------------------------
                // blank
                // ----------------------------------------
                if (quoteRowKind == RenderState.QuoteRowKind.BLANK) {

                    boolean isQuoteDecorCol = c >= startCol && c < startCol + quoteDepth;

                    cell.setCellStyle(
                            isQuoteDecorCol ? styles.blockQuoteBlankLeftStyle : styles.blockQuoteBlankBodyStyle);

                    continue;
                }

                // ----------------------------------------
                // normal / heading / list
                // ----------------------------------------
                boolean isQuoteDecorCol = c >= startCol && c < startCol + quoteDepth;

                if (isQuoteDecorCol) {
                    cell.setCellStyle(styles.blockQuoteLeftStyle);

                } else {
                    cell.setCellStyle(resolveBlockQuoteContentStyle(cell.getCellStyle(), styles));
                }
            }
        }
    }

    private static CellStyle resolveBlockQuoteContentStyle(CellStyle currentStyle, MdStyle styles) {

        int styleIndex = currentStyle.getIndex();

        if (styleIndex == styles.heading1Style.getIndex() || styleIndex == styles.blockQuoteHeading1Style.getIndex()) {
            return styles.blockQuoteHeading1Style;
        }

        if (styleIndex == styles.heading2Style.getIndex() || styleIndex == styles.blockQuoteHeading2Style.getIndex()) {
            return styles.blockQuoteHeading2Style;
        }

        if (styleIndex == styles.heading3Style.getIndex() || styleIndex == styles.blockQuoteHeading3Style.getIndex()) {
            return styles.blockQuoteHeading3Style;
        }

        if (styleIndex == styles.heading4Style.getIndex() || styleIndex == styles.blockQuoteHeading4Style.getIndex()) {
            return styles.blockQuoteHeading4Style;
        }

        return styles.blockQuoteBodyStyle;
    }
}