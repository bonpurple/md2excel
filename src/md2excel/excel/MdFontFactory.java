package md2excel.excel;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

final class MdFontFactory {

    private final XSSFWorkbook workbook;

    MdFontFactory(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    Font create(String fontName, int fontSize, boolean bold, boolean italic) {

        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeightInPoints((short) fontSize);
        font.setBold(bold);
        font.setItalic(italic);

        return font;
    }
}