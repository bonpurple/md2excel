package md2excel.app;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import md2excel.config.MdFontSettings;
import md2excel.config.MdSheetSettings;
import md2excel.excel.ExcelCellUtil;
import md2excel.excel.MdStyleCatalog;
import md2excel.render.MarkdownRenderer;
import md2excel.render.RenderContext;
import md2excel.render.SheetColumnLayout;

/**
 * Markdownの行をExcel Workbookへ描画する。
 *
 * ファイル入出力は担当しない。
 */
final class MarkdownWorkbookRenderer {

    static final String SHEET_NAME = "spec";

    // POI内部幅。Excel表示上は約2.5
    private static final int DEFAULT_COLUMN_WIDTH = (int) (3.125 * 256);

    private static final int START_ROW_INDEX = 1;

    private static final int LEFT_MARGIN_COLUMN_COUNT = 1;
    private static final int RIGHT_MARGIN_COLUMN_COUNT = 1;

    void render(Iterator<String> lines, XSSFWorkbook workbook, MdFontSettings fontSettings,
            MdSheetSettings sheetSettings) {

        if (lines == null) {
            throw new IllegalArgumentException("lines must not be null");
        }

        if (workbook == null) {
            throw new IllegalArgumentException("workbook must not be null");
        }

        if (fontSettings == null) {
            throw new IllegalArgumentException("fontSettings must not be null");
        }

        if (sheetSettings == null) {
            throw new IllegalArgumentException("sheetSettings must not be null");
        }

        Sheet sheet = workbook.createSheet(SHEET_NAME);

        initializeSheet(sheet);

        MdStyleCatalog styles = createStyles(workbook, fontSettings, sheetSettings);

        SheetColumnLayout columnLayout = createColumnLayout(sheetSettings);

        initializeColumns(sheet, styles, columnLayout.getTotalColumnCount());

        prefillRows(sheet, styles, columnLayout.getTotalColumnCount());

        RenderContext context = new RenderContext(workbook, sheet, styles, columnLayout, START_ROW_INDEX);

        MarkdownRenderer.render(lines, context);
    }

    private static MdStyleCatalog createStyles(XSSFWorkbook workbook, MdFontSettings fontSettings,
            MdSheetSettings sheetSettings) {

        return new MdStyleCatalog(workbook, fontSettings.getFontName(), fontSettings.getH1Size(),
                fontSettings.getH2Size(), fontSettings.getH3Size(), fontSettings.getNormalSize(),
                sheetSettings.getVerticalAlignment());
    }

    private static SheetColumnLayout createColumnLayout(MdSheetSettings sheetSettings) {

        return new SheetColumnLayout(sheetSettings.getTotalColumnCount(), LEFT_MARGIN_COLUMN_COUNT,
                RIGHT_MARGIN_COLUMN_COUNT);
    }

    private static void initializeSheet(Sheet sheet) {
        sheet.setDisplayGridlines(false);
        sheet.setPrintGridlines(false);
    }

    private static void initializeColumns(Sheet sheet, MdStyleCatalog styles, int totalColumnCount) {

        for (int column = 0; column < totalColumnCount; column++) {

            sheet.setColumnWidth(column, DEFAULT_COLUMN_WIDTH);

            sheet.setDefaultColumnStyle(column, styles.normalStyle);
        }
    }

    private static void prefillRows(Sheet sheet, MdStyleCatalog styles, int totalColumnCount) {

        for (int rowIndex = 0; rowIndex < START_ROW_INDEX; rowIndex++) {

            Row row = sheet.getRow(rowIndex);

            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            row.setRowStyle(styles.blankRowStyle);

            for (int column = 0; column < totalColumnCount; column++) {

                Cell cell = ExcelCellUtil.getOrCreateCell(row, column);

                cell.setBlank();
                cell.setCellStyle(styles.blankRowStyle);
            }
        }
    }
}