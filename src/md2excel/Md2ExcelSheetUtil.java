package md2excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public final class Md2ExcelSheetUtil {
    private Md2ExcelSheetUtil() {
    }

    public static void createHorizontalRuleRow(Sheet sheet, Row row, CellStyle style, int mergeCols) {
        for (int c = 0; c < mergeCols; c++) {
            Cell cell = row.createCell(c);
            cell.setCellStyle(style);
        }
    }
}
