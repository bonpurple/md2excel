package md2excel.config;

import java.nio.file.Path;

/**
 * MarkdownからExcelへの変換設定。
 */
public final class Md2ExcelConfig {

    private final Path inputPath;
    private final Path outputPath;

    private final MdFontSettings fontSettings;
    private final MdSheetSettings sheetSettings;

    public Md2ExcelConfig(Path inputPath, Path outputPath, MdFontSettings fontSettings, MdSheetSettings sheetSettings) {

        this.inputPath = requirePath(inputPath, "inputPath");

        this.outputPath = requirePath(outputPath, "outputPath");

        this.fontSettings = requireValue(fontSettings, "fontSettings");

        this.sheetSettings = requireValue(sheetSettings, "sheetSettings");
    }

    public Path getInputPath() {
        return inputPath;
    }

    public Path getOutputPath() {
        return outputPath;
    }

    public MdFontSettings getFontSettings() {
        return fontSettings;
    }

    public MdSheetSettings getSheetSettings() {
        return sheetSettings;
    }

    private static Path requirePath(Path value, String name) {

        if (value == null) {
            throw new IllegalArgumentException(name + " must not be null");
        }

        if (value.toString().isEmpty()) {
            throw new IllegalArgumentException(name + " must not be empty");
        }

        return value;
    }

    private static <T> T requireValue(T value, String name) {

        if (value == null) {
            throw new IllegalArgumentException(name + " must not be null");
        }

        return value;
    }
}