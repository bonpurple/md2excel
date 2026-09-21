package md2excel.app;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import md2excel.config.Md2ExcelConfig;
import md2excel.excel.ExcelCellUtil;
import md2excel.excel.MdStyleCatalog;
import md2excel.render.MarkdownRenderer;
import md2excel.render.RenderContext;

public final class MarkdownToExcelConverter {

    // POI内部幅。Excel表示上は約2.5
    private static final int DEFAULT_COLUMN_WIDTH = (int) (3.125 * 256);

    private static final int START_ROW_INDEX = 1;
    private static final int START_COL_INDEX = 1;

    public void convert(Md2ExcelConfig config) throws IOException {
        Path markdownPath = Paths.get(config.inPath);
        Path excelPath = Paths.get(config.outPath);

        try (Stream<String> lines = Files.lines(markdownPath, StandardCharsets.UTF_8);
                XSSFWorkbook workbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("spec");
            initializeSheet(sheet);

            MdStyleCatalog styles = new MdStyleCatalog(workbook, config.fontName, config.h1Size, config.h2Size,
                    config.h3Size, config.normalSize, config.vAlign);

            initializeColumns(sheet, styles, config.sheetColumnCount);

            prefillRows(sheet, styles, config.sheetColumnCount);

            RenderContext context = new RenderContext(workbook, sheet, styles, config.sheetColumnCount, START_ROW_INDEX,
                    START_COL_INDEX);

            MarkdownRenderer.render(lines.iterator(), context);

            try (OutputStream output = Files.newOutputStream(excelPath)) {
                workbook.write(output);
            }
        }
    }

    private static void initializeSheet(Sheet sheet) {
        sheet.setDisplayGridlines(false);
        sheet.setPrintGridlines(false);
    }

    private static void initializeColumns(Sheet sheet, MdStyleCatalog styles, int sheetColumnCount) {

        for (int column = 0; column < sheetColumnCount; column++) {

            sheet.setColumnWidth(column, DEFAULT_COLUMN_WIDTH);

            sheet.setDefaultColumnStyle(column, styles.normalStyle);
        }
    }

    private static void prefillRows(Sheet sheet, MdStyleCatalog styles, int sheetColumnCount) {

        for (int rowIndex = 0; rowIndex < START_ROW_INDEX; rowIndex++) {

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            row.setRowStyle(styles.blankRowStyle);

            for (int column = 0; column < sheetColumnCount; column++) {

                Cell cell = ExcelCellUtil.getOrCreateCell(row, column);

                cell.setBlank();
                cell.setCellStyle(styles.blankRowStyle);
            }
        }
    }
}