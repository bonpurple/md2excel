package md2excel.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public final class Md2ExcelSheetUtil {
    private Md2ExcelSheetUtil() {
    }

    public static void createHorizontalRuleRow(Sheet sheet, Row row, CellStyle style, int startCol,
            int endColExclusive) {
        for (int c = startCol; c < endColExclusive; c++) {
            Cell cell = ExcelCellUtil.getOrCreateCell(row, c);
            cell.setCellStyle(style);
        }
    }
}