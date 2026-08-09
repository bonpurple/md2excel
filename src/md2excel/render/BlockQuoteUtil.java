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
            st.blankBlockQuoteRows.clear();
            st.horizontalRuleBlockQuoteRows.clear();
            st.tableBlockQuoteRows.clear();
            return;
        }
        if (st.blockQuoteFirstRow < 0 || st.blockQuoteLastRow < 0) {
            st.blankBlockQuoteRows.clear();
            st.horizontalRuleBlockQuoteRows.clear();
            st.tableBlockQuoteRows.clear();
            return;
        }

        applyBlockQuoteStyle(sheet, styles, st, st.blockQuoteFirstRow, st.blockQuoteLastRow, st.blockQuoteCol,
                st.lastColIndex);

        st.inBlockQuote = false;
        st.blockQuoteFirstRow = -1;
        st.blockQuoteLastRow = -1;
        st.blockQuoteCellRow = -1;
        st.blockQuoteCellCol = -1;
        st.blankBlockQuoteRows.clear();
        st.tableBlockQuoteRows.clear();
    }

    private static void applyBlockQuoteStyle(Sheet sheet, MdStyle styles, RenderState st, int firstRow, int lastRow,
            int startCol, int lastColIndex) {

        int fillEndCol = Math.max(startCol, lastColIndex);

        for (int r = firstRow; r <= lastRow; r++) {
            Row rowObj = sheet.getRow(r);
            if (rowObj == null)
                continue;

            boolean blankQuoteRow = st.blankBlockQuoteRows.contains(r);

            boolean horizontalRuleQuoteRow = st.horizontalRuleBlockQuoteRows.contains(r);

            boolean tableQuoteRow = st.tableBlockQuoteRows.contains(r);

            for (int c = startCol; c <= fillEndCol; c++) {
                Cell cell = rowObj.getCell(c);
                if (cell == null) {
                    cell = rowObj.createCell(c);
                    cell.setBlank();
                }

                if (horizontalRuleQuoteRow) {
                    boolean isLeft = (c == startCol);

                    cell.setCellStyle(isLeft ? styles.blockQuoteHorizontalRuleLeftStyle
                            : styles.blockQuoteHorizontalRuleBodyStyle);

                    continue;
                }

                if (tableQuoteRow) {
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

                boolean isLeft = (c == startCol);

                if (blankQuoteRow) {
                    cell.setCellStyle(isLeft ? styles.blockQuoteBlankLeftStyle : styles.blockQuoteBlankBodyStyle);

                } else if (isLeft) {
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