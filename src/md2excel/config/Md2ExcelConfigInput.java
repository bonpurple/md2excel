package md2excel.config;

import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 設定フォームの入力値から変換設定を作る。
 */
final class Md2ExcelConfigInput {

    private Md2ExcelConfigInput() {
    }

    static Md2ExcelConfig create(String inputText, String fontName, VerticalAlignment alignment, int h1Size, int h2Size,
            int h3Size, int normalSize, int totalColumnCount) {

        String trimmedInput = inputText.trim();

        if (trimmedInput.isEmpty()) {
            throw new IllegalArgumentException("Markdownファイルを選択してください。");
        }

        Path inputPath = Paths.get(trimmedInput).toAbsolutePath();
        Path outputPath = replaceExtension(inputPath, ".xlsx");

        if (fontName == null || fontName.trim().isEmpty()) {
            throw new IllegalArgumentException("フォントを選択してください。");
        }

        if (alignment == null) {
            throw new IllegalArgumentException("セルの縦位置を選択してください。");
        }

        MdFontSettings fontSettings = new MdFontSettings(fontName, h1Size, h2Size, h3Size, normalSize);
        MdSheetSettings sheetSettings = new MdSheetSettings(totalColumnCount, alignment);

        return new Md2ExcelConfig(inputPath, outputPath, fontSettings, sheetSettings);
    }

    static Path replaceExtension(Path path, String newExtension) {

        if (path == null) {
            throw new IllegalArgumentException("path must not be null");
        }

        if (newExtension == null || newExtension.trim().isEmpty()) {
            throw new IllegalArgumentException("newExtension must not be empty");
        }

        Path fileNamePath = path.getFileName();

        if (fileNamePath == null) {
            throw new IllegalArgumentException("path must have a file name: " + path);
        }

        String fileName = fileNamePath.toString();
        int dot = fileName.lastIndexOf('.');
        String outputFileName = dot > 0 ? fileName.substring(0, dot) + newExtension : fileName + newExtension;
        Path parent = path.getParent();

        return parent == null ? Paths.get(outputFileName) : parent.resolve(outputFileName);
    }
}
