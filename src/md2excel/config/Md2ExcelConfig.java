package md2excel.config;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public final class Md2ExcelConfig {

    // B列開始と左右の余白を確保するため、最低3列（A～C）とする
    public static final int MIN_SHEET_COLUMN_COUNT = 3;
    public static final int MAX_SHEET_COLUMN_COUNT = SpreadsheetVersion.EXCEL2007.getMaxColumns();

    public static final int MIN_FONT_SIZE = 5;
    public static final int MAX_FONT_SIZE = 72;

    public final String inPath;
    public final String outPath;
    public final int sheetColumnCount;
    public final String fontName;
    public final int h1Size;
    public final int h2Size;
    public final int h3Size;
    public final int normalSize;
    public final VerticalAlignment vAlign;

    public Md2ExcelConfig(String inPath, String outPath, int sheetColumnCount, String fontName, int h1Size, int h2Size,
            int h3Size, int normalSize, VerticalAlignment vAlign) {

        this.inPath = requireText(inPath, "inPath");
        this.outPath = requireText(outPath, "outPath");
        this.fontName = requireText(fontName, "fontName");

        validateRange(sheetColumnCount, MIN_SHEET_COLUMN_COUNT, MAX_SHEET_COLUMN_COUNT, "sheetColumnCount");

        validateRange(h1Size, MIN_FONT_SIZE, MAX_FONT_SIZE, "h1Size");

        validateRange(h2Size, MIN_FONT_SIZE, MAX_FONT_SIZE, "h2Size");

        validateRange(h3Size, MIN_FONT_SIZE, MAX_FONT_SIZE, "h3Size");

        validateRange(normalSize, MIN_FONT_SIZE, MAX_FONT_SIZE, "normalSize");

        if (vAlign == null) {
            throw new IllegalArgumentException("vAlign must not be null");
        }

        this.sheetColumnCount = sheetColumnCount;
        this.h1Size = h1Size;
        this.h2Size = h2Size;
        this.h3Size = h3Size;
        this.normalSize = normalSize;
        this.vAlign = vAlign;
    }

    private static String requireText(String value, String name) {

        if (value == null || value.trim().isEmpty()) {
            throw new IllegalArgumentException(name + " must not be empty");
        }

        return value.trim();
    }

    private static void validateRange(int value, int min, int max, String name) {

        if (value < min || value > max) {
            throw new IllegalArgumentException(name + " must be between " + min + " and " + max + ": " + value);
        }
    }
}