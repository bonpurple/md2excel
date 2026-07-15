package md2excel.app;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import md2excel.config.Md2ExcelConfig;
import md2excel.excel.MdStyle;
import md2excel.render.MarkdownRenderer;
import md2excel.render.RenderContext;

public class MarkdownToExcel {

    // POI内部幅。Excel表示上は約 2.5
    private static final int DEFAULT_COLUMN_WIDTH = (int) (3.125 * 256);

    private static final int START_ROW_INDEX = 1; // Excelの2行目
    private static final int START_COL_INDEX = 1; // ExcelのB列

    public static void main(String[] args) throws Exception {
        Md2ExcelConfig cfg = Md2ExcelConfig.load(args);
        if (cfg == null) {
            System.out.println("キャンセルされました。処理を終了します。");
            return;
        }

        Path mdPath = Paths.get(cfg.inPath);
        Path xlsxPath = Paths.get(cfg.outPath);

        // 全読みをやめて逐次読み（Stream）にする
        try (Stream<String> lines = Files.lines(mdPath, StandardCharsets.UTF_8);
                Workbook workbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("spec");

            sheet.setDisplayGridlines(false);
            sheet.setPrintGridlines(false);

            MdStyle styles = new MdStyle(workbook, cfg.fontName, cfg.h1Size, cfg.h2Size, cfg.h3Size, cfg.normalSize,
                    cfg.vAlign);

            for (int c = 0; c < cfg.mergeCols; c++) {
                sheet.setColumnWidth(c, DEFAULT_COLUMN_WIDTH);
                sheet.setDefaultColumnStyle(c, styles.normalStyle);
            }

            // 開始位置より前の空行にも blankRowStyle を適用する
            int prefillColCount = cfg.mergeCols;

            for (int r = 0; r < START_ROW_INDEX; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    row = sheet.createRow(r);
                }
                row.setRowStyle(styles.blankRowStyle);

                for (int c = 0; c < prefillColCount; c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null) {
                        cell = row.createCell(c);
                    }
                    cell.setBlank();
                    cell.setCellStyle(styles.blankRowStyle);
                }
            }

            RenderContext ctx = new RenderContext(workbook, sheet, styles, cfg.mergeCols, START_ROW_INDEX,
                    START_COL_INDEX);

            MarkdownRenderer.render(lines.iterator(), ctx);

            try (OutputStream os = Files.newOutputStream(xlsxPath)) {
                workbook.write(os);
            }

            System.out.println("生成完了: " + xlsxPath.toAbsolutePath());
            JOptionPane.showMessageDialog(null, "Excel ファイルを生成しました。\n" + xlsxPath.toAbsolutePath(), "完了",
                    JOptionPane.INFORMATION_MESSAGE);
        }
    }
}