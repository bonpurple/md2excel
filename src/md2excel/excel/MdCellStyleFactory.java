package md2excel.excel;

import java.awt.Color;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

final class MdCellStyleFactory {

    private final XSSFWorkbook workbook;

    MdCellStyleFactory(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    XSSFCellStyle createBase(VerticalAlignment verticalAlignment) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setWrapText(false);
        style.setVerticalAlignment(verticalAlignment);
        return style;
    }

    XSSFCellStyle cloneOf(CellStyle baseStyle) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(baseStyle);
        return style;
    }

    XSSFCellStyle cloneWithFont(CellStyle baseStyle, Font font) {

        XSSFCellStyle style = cloneOf(baseStyle);
        style.setFont(font);
        return style;
    }

    XSSFCellStyle cloneWithFill(CellStyle baseStyle, XSSFColor background) {

        XSSFCellStyle style = cloneOf(baseStyle);
        applyFill(style, background);
        return style;
    }

    void applyFill(XSSFCellStyle style, XSSFColor background) {

        if (background == null) {
            return;
        }

        style.setFillForegroundColor(background);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    XSSFColor color(Color color) {
        return new XSSFColor(color, null);
    }

    void clearBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.NONE);
    }
}