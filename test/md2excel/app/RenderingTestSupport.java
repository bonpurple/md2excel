package md2excel.app;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import md2excel.config.MdFontSettings;
import md2excel.config.MdSheetSettings;

final class RenderingTestSupport {

    private RenderingTestSupport() {
    }

    static Sheet render(XSSFWorkbook workbook, String... lines) {
        return render(workbook, 40, lines);
    }

    static Sheet render(XSSFWorkbook workbook, int totalColumnCount, String... lines) {
        new MarkdownWorkbookRenderer().render(Arrays.asList(lines).iterator(), workbook,
                new MdFontSettings("Meiryo", 16, 14, 12, 11),
                new MdSheetSettings(totalColumnCount, VerticalAlignment.BOTTOM));
        return workbook.getSheet("spec");
    }

    static List<String> textCells(Sheet sheet) {
        List<String> actual = new ArrayList<String>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && !cell.getStringCellValue().isEmpty()) {
                    actual.add(cell.getAddress().toString() + "=" + cell.getStringCellValue());
                }
            }
        }
        return actual;
    }
}
