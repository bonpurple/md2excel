package md2excel.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public final class ExcelCellUtil {

    private ExcelCellUtil() {
    }

    public static Cell getOrCreateCell(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);

        if (cell == null) {
            cell = row.createCell(colIndex);
        }

        return cell;
    }
}