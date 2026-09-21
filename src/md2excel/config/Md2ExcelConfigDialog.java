package md2excel.config;

import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.VerticalAlignment;

public final class Md2ExcelConfigDialog {

    private static final String DEFAULT_FONT_NAME = "游ゴシック";
    private static final int DEFAULT_H1_FONT_SIZE = 16;
    private static final int DEFAULT_H2_FONT_SIZE = 14;
    private static final int DEFAULT_H3_FONT_SIZE = 12;
    private static final int DEFAULT_NORMAL_FONT_SIZE = 11;
    private static final int DEFAULT_SHEET_COLUMN_COUNT = 40;

    private Md2ExcelConfigDialog() {
    }

    public static Md2ExcelConfig load() {
        File mdFile = chooseMarkdownFile();
        if (mdFile == null) {
            return null;
        }

        String inPath = mdFile.getAbsolutePath();
        String outPath = replaceExtension(inPath, ".xlsx");

        int sheetColumnCount = askSheetColumnCount();

        String[] fontCandidates = { "游ゴシック", "Yu Gothic UI", "ＭＳ Ｐゴシック", "ＭＳ ゴシック", "Meiryo", "Meiryo UI" };

        Object selectedFont = JOptionPane.showInputDialog(null, "フォントを選択してください（キャンセルで既定のフォント）。", "フォント選択",
                JOptionPane.QUESTION_MESSAGE, null, fontCandidates, DEFAULT_FONT_NAME);

        String fontName = selectedFont == null ? DEFAULT_FONT_NAME : selectedFont.toString().trim();

        String[] valignOptions = { "上揃え", "上下中央揃え", "下揃え" };

        Object selectedAlign = JOptionPane.showInputDialog(null, "セルの縦方向の配置を選択してください。", "縦位置",
                JOptionPane.QUESTION_MESSAGE, null, valignOptions, "下揃え");

        VerticalAlignment verticalAlignment = selectedAlign == null ? VerticalAlignment.BOTTOM
                : toVerticalAlignment(selectedAlign.toString());

        int h1Size = askFontSize("# 見出しのフォントサイズ (pt) を入力してください。", DEFAULT_H1_FONT_SIZE);

        int h2Size = askFontSize("## 見出しのフォントサイズ (pt) を入力してください。", DEFAULT_H2_FONT_SIZE);

        int h3Size = askFontSize("### 見出しのフォントサイズ (pt) を入力してください。", DEFAULT_H3_FONT_SIZE);

        int normalSize = askFontSize("通常テキストのフォントサイズ (pt) を入力してください。", DEFAULT_NORMAL_FONT_SIZE);

        return new Md2ExcelConfig(inPath, outPath, sheetColumnCount, fontName, h1Size, h2Size, h3Size, normalSize,
                verticalAlignment);
    }

    private static int askSheetColumnCount() {
        while (true) {
            String input = JOptionPane
                    .showInputDialog(null,
                            "シート左端（A列）基準の列数を入力してください。\n" + Md2ExcelConfig.MIN_SHEET_COLUMN_COUNT + "～"
                                    + Md2ExcelConfig.MAX_SHEET_COLUMN_COUNT,
                            Integer.toString(DEFAULT_SHEET_COLUMN_COUNT));

            // キャンセルまたは空入力は既定値
            if (input == null || input.trim().isEmpty()) {
                return DEFAULT_SHEET_COLUMN_COUNT;
            }

            try {
                int value = Integer.parseInt(input.trim());

                if (value >= Md2ExcelConfig.MIN_SHEET_COLUMN_COUNT && value <= Md2ExcelConfig.MAX_SHEET_COLUMN_COUNT) {
                    return value;
                }
            } catch (NumberFormatException e) {
                // 下の警告を表示して再入力
            }

            JOptionPane.showMessageDialog(null, Md2ExcelConfig.MIN_SHEET_COLUMN_COUNT + "～"
                    + Md2ExcelConfig.MAX_SHEET_COLUMN_COUNT + "の整数を入力してください。", "入力エラー", JOptionPane.WARNING_MESSAGE);
        }
    }

    private static int askFontSize(String message, int defaultSize) {

        while (true) {
            String input = JOptionPane.showInputDialog(null,
                    message + "\n" + Md2ExcelConfig.MIN_FONT_SIZE + "～" + Md2ExcelConfig.MAX_FONT_SIZE,
                    Integer.toString(defaultSize));

            // キャンセルまたは空入力は既定値
            if (input == null || input.trim().isEmpty()) {
                return defaultSize;
            }

            try {
                int value = Integer.parseInt(input.trim());

                if (value >= Md2ExcelConfig.MIN_FONT_SIZE && value <= Md2ExcelConfig.MAX_FONT_SIZE) {
                    return value;
                }
            } catch (NumberFormatException e) {
                // 下の警告を表示して再入力
            }

            JOptionPane.showMessageDialog(null,
                    Md2ExcelConfig.MIN_FONT_SIZE + "～" + Md2ExcelConfig.MAX_FONT_SIZE + "の整数を入力してください。", "入力エラー",
                    JOptionPane.WARNING_MESSAGE);
        }
    }

    private static File chooseMarkdownFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Markdown ファイルを選択してください");
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

        int result = fileChooser.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            return fileChooser.getSelectedFile();
        }

        return null;
    }

    private static VerticalAlignment toVerticalAlignment(String label) {
        switch (label) {
        case "上揃え":
            return VerticalAlignment.TOP;

        case "下揃え":
            return VerticalAlignment.BOTTOM;

        case "上下中央揃え":
        default:
            return VerticalAlignment.CENTER;
        }
    }

    private static String replaceExtension(String path, String newExtension) {

        File file = new File(path);
        String fileName = file.getName();

        int dot = fileName.lastIndexOf('.');
        String outputFileName = dot > 0 ? fileName.substring(0, dot) + newExtension : fileName + newExtension;

        File parent = file.getParentFile();
        return parent == null ? outputFileName : new File(parent, outputFileName).getPath();
    }
}