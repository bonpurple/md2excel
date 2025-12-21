package md2excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public final class BlockQuoteUtil {
    private BlockQuoteUtil() {
    }

    public static void closeBlockQuoteIfOpen(Sheet sheet, MdStyle styles, RenderState st) {
        if (!st.inBlockQuote)
            return;
        if (st.blockQuoteFirstRow < 0 || st.blockQuoteLastRow < 0)
            return;

        applyBlockQuoteStyle(sheet, styles, st.blockQuoteFirstRow, st.blockQuoteLastRow, st.blockQuoteCol,
                st.lastColIndex);

        st.inBlockQuote = false;
        st.blockQuoteFirstRow = -1;
        st.blockQuoteLastRow = -1;
        st.blockQuoteCellRow = -1;
        st.blockQuoteCellCol = -1;
    }

    private static void applyBlockQuoteStyle(Sheet sheet, MdStyle styles, int firstRow, int lastRow, int startCol,
            int lastColIndex) {

        int fillEndCol = Math.max(startCol, lastColIndex);

        for (int r = firstRow; r <= lastRow; r++) {
            Row rowObj = sheet.getRow(r);
            if (rowObj == null)
                continue;

            for (int c = startCol; c <= fillEndCol; c++) {
                Cell cell = rowObj.getCell(c);
                if (cell == null) {
                    cell = rowObj.createCell(c);
                    cell.setBlank();
                }

                boolean isLeft = (c == startCol);
                cell.setCellStyle(isLeft ? styles.blockQuoteLeftStyle : styles.blockQuoteBodyStyle);
            }
        }
    }
}
