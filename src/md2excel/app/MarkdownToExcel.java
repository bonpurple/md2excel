package md2excel.app;

import java.nio.file.Path;
import java.nio.file.Paths;

import javax.swing.JOptionPane;

import md2excel.config.Md2ExcelConfig;
import md2excel.config.Md2ExcelConfigDialog;

public final class MarkdownToExcel {

    private MarkdownToExcel() {
    }

    public static void main(String[] args) throws Exception {
        Md2ExcelConfig config = Md2ExcelConfigDialog.load();

        if (config == null) {
            System.out.println("キャンセルされました。処理を終了します。");
            return;
        }

        MarkdownToExcelConverter converter = new MarkdownToExcelConverter();

        converter.convert(config);

        Path outputPath = Paths.get(config.outPath).toAbsolutePath();

        System.out.println("生成完了: " + outputPath);

        JOptionPane.showMessageDialog(null, "Excel ファイルを生成しました。\n" + outputPath, "完了", JOptionPane.INFORMATION_MESSAGE);
    }
}