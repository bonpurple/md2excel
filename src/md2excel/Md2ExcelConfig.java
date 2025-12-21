package md2excel;

import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.VerticalAlignment;

public final class Md2ExcelConfig {
    public final String inPath;
    public final String outPath;
    public final int mergeCols;
    public final String fontName;
    public final int h1Size;
    public final int h2Size;
    public final int h3Size;
    public final int normalSize;
    public final VerticalAlignment vAlign;

    // 既定値
    private static final String DEFAULT_FONT_NAME = "游ゴシック";
    private static final int DEFAULT_H1_FONT_SIZE = 16;
    private static final int DEFAULT_H2_FONT_SIZE = 14;
    private static final int DEFAULT_H3_FONT_SIZE = 12;
    private static final int DEFAULT_NORMAL_FONT_SIZE = 10;
    private static final int DEFAULT_MERGE_COLS = 40;

    private Md2ExcelConfig(String in, String out, int mergeCols, String fontName, int h1, int h2, int h3, int normal,
            VerticalAlignment vAlign) {
        this.inPath = in;
        this.outPath = out;
        this.mergeCols = mergeCols;
        this.fontName = fontName;
        this.h1Size = h1;
        this.h2Size = h2;
        this.h3Size = h3;
        this.normalSize = normal;
        this.vAlign = vAlign;
    }

    public static Md2ExcelConfig load(String[] args) {
        if (args != null && args.length >= 1) {
            String in = args[0];
            String out = (args.length >= 2) ? args[1] : replaceExtension(args[0], ".xlsx");
            int mergeCols = (args.length >= 3) ? parseIntOrDefault(args[2], DEFAULT_MERGE_COLS) : DEFAULT_MERGE_COLS;
            String fontName = (args.length >= 4 && args[3] != null && !args[3].trim().isEmpty()) ? args[3].trim()
                    : DEFAULT_FONT_NAME;

            return new Md2ExcelConfig(in, out, mergeCols, fontName, DEFAULT_H1_FONT_SIZE, DEFAULT_H2_FONT_SIZE,
                    DEFAULT_H3_FONT_SIZE, DEFAULT_NORMAL_FONT_SIZE, VerticalAlignment.CENTER);
        }

        // 引数なし → ダイアログ
        File mdFile = chooseMarkdownFile();
        if (mdFile == null) {
            return null;
        }
        String in = mdFile.getAbsolutePath();
        String out = replaceExtension(in, ".xlsx");

        String inputCols = JOptionPane.showInputDialog(null, "1行分として扱う列数（MERGE_LAST_COL）を入力してください。", "40");
        int mergeCols = parseIntOrDefault(inputCols, DEFAULT_MERGE_COLS);

        // フォント選択
        String[] fontCandidates = { "游ゴシック", "Yu Gothic UI", "ＭＳ Ｐゴシック", "ＭＳ ゴシック", "Meiryo", "Meiryo UI" };
        Object selectedFont = JOptionPane.showInputDialog(null, "フォントを選択してください（キャンセルで既定のフォント）。", "フォント選択",
                JOptionPane.QUESTION_MESSAGE, null, fontCandidates, DEFAULT_FONT_NAME);
        String fontName = (selectedFont == null) ? DEFAULT_FONT_NAME : selectedFont.toString().trim();

        // 縦位置
        String[] valignOptions = { "上揃え", "上下中央揃え", "下揃え" };
        Object selectedAlign = JOptionPane.showInputDialog(null, "セルの縦方向の配置を選択してください。", "縦位置",
                JOptionPane.QUESTION_MESSAGE, null, valignOptions, "上下中央揃え");
        VerticalAlignment vAlign = (selectedAlign == null) ? VerticalAlignment.CENTER
                : toVerticalAlignment(selectedAlign.toString());

        // サイズ
        int h1 = parseFontSize(JOptionPane.showInputDialog(null, "# 見出しのフォントサイズ (pt) を入力してください。",
                Integer.toString(DEFAULT_H1_FONT_SIZE)), DEFAULT_H1_FONT_SIZE);

        int h2 = parseFontSize(JOptionPane.showInputDialog(null, "## 見出しのフォントサイズ (pt) を入力してください。",
                Integer.toString(DEFAULT_H2_FONT_SIZE)), DEFAULT_H2_FONT_SIZE);

        int h3 = parseFontSize(JOptionPane.showInputDialog(null, "### 見出しのフォントサイズ (pt) を入力してください。",
                Integer.toString(DEFAULT_H3_FONT_SIZE)), DEFAULT_H3_FONT_SIZE);

        int normal = parseFontSize(JOptionPane.showInputDialog(null, "通常テキストのフォントサイズ (pt) を入力してください。",
                Integer.toString(DEFAULT_NORMAL_FONT_SIZE)), DEFAULT_NORMAL_FONT_SIZE);

        return new Md2ExcelConfig(in, out, mergeCols, fontName, h1, h2, h3, normal, vAlign);
    }

    private static File chooseMarkdownFile() {
        JFileChooser fc = new JFileChooser();
        fc.setDialogTitle("Markdown ファイルを選択してください");
        fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
        int result = fc.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            return fc.getSelectedFile();
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

    private static int parseFontSize(String s, int defaultSize) {
        if (s == null || s.trim().isEmpty()) {
            return defaultSize;
        }
        try {
            int v = Integer.parseInt(s.trim());
            if (v < 5 || v > 72) {
                return defaultSize;
            }
            return v;
        } catch (NumberFormatException e) {
            return defaultSize;
        }
    }

    private static int parseIntOrDefault(String s, int defaultValue) {
        if (s == null || s.trim().isEmpty()) {
            return defaultValue;
        }
        try {
            int v = Integer.parseInt(s.trim());
            return v > 0 ? v : defaultValue;
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    private static String replaceExtension(String path, String newExt) {
        int dot = path.lastIndexOf('.');
        if (dot == -1) {
            return path + newExt;
        }
        return path.substring(0, dot) + newExt;
    }
}
