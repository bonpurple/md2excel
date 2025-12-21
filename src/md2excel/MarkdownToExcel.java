package md2excel;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MarkdownToExcel {

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
                sheet.setColumnWidth(c, 3 * 256);
                sheet.setDefaultColumnStyle(c, styles.normalStyle);
            }

            RenderContext ctx = new RenderContext(workbook, sheet, styles, cfg.mergeCols);

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
