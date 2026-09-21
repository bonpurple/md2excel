package md2excel.config;

import java.util.Objects;

/**
 * Markdownを描画するときのフォント設定。
 */
public final class MdFontSettings {

    public static final int MIN_FONT_SIZE = 5;
    public static final int MAX_FONT_SIZE = 72;

    private final String fontName;
    private final int h1Size;
    private final int h2Size;
    private final int h3Size;
    private final int normalSize;

    public MdFontSettings(String fontName, int h1Size, int h2Size, int h3Size, int normalSize) {

        this.fontName = requireText(fontName, "fontName");

        this.h1Size = validateFontSize(h1Size, "h1Size");

        this.h2Size = validateFontSize(h2Size, "h2Size");

        this.h3Size = validateFontSize(h3Size, "h3Size");

        this.normalSize = validateFontSize(normalSize, "normalSize");
    }

    public String getFontName() {
        return fontName;
    }

    public int getH1Size() {
        return h1Size;
    }

    public int getH2Size() {
        return h2Size;
    }

    public int getH3Size() {
        return h3Size;
    }

    public int getNormalSize() {
        return normalSize;
    }

    private static String requireText(String value, String name) {

        if (value == null || value.trim().isEmpty()) {
            throw new IllegalArgumentException(name + " must not be empty");
        }

        return value.trim();
    }

    private static int validateFontSize(int value, String name) {

        if (value < MIN_FONT_SIZE || value > MAX_FONT_SIZE) {

            throw new IllegalArgumentException(
                    name + " must be between " + MIN_FONT_SIZE + " and " + MAX_FONT_SIZE + ": " + value);
        }

        return value;
    }

    @Override
    public boolean equals(Object other) {
        if (this == other) {
            return true;
        }

        if (!(other instanceof MdFontSettings)) {
            return false;
        }

        MdFontSettings that = (MdFontSettings) other;

        return h1Size == that.h1Size && h2Size == that.h2Size && h3Size == that.h3Size && normalSize == that.normalSize
                && fontName.equals(that.fontName);
    }

    @Override
    public int hashCode() {
        return Objects.hash(fontName, Integer.valueOf(h1Size), Integer.valueOf(h2Size), Integer.valueOf(h3Size),
                Integer.valueOf(normalSize));
    }

    @Override
    public String toString() {
        return "MdFontSettings{" + "fontName='" + fontName + '\'' + ", h1Size=" + h1Size + ", h2Size=" + h2Size
                + ", h3Size=" + h3Size + ", normalSize=" + normalSize + '}';
    }
}