package md2excel.app;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Workbookを一時ファイルへ書き込み、 成功後に最終出力ファイルへ置換する。
 */
final class AtomicWorkbookWriter {

    private static final String TEMP_FILE_PREFIX = ".md2excel-";

    private static final String TEMP_FILE_SUFFIX = ".xlsx.tmp";

    void write(XSSFWorkbook workbook, Path outputPath) throws IOException {

        if (workbook == null) {
            throw new IllegalArgumentException("workbook must not be null");
        }

        if (outputPath == null) {
            throw new IllegalArgumentException("outputPath must not be null");
        }

        Path absoluteOutputPath = outputPath.toAbsolutePath();

        Path outputDirectory = absoluteOutputPath.getParent();

        if (outputDirectory == null) {
            throw new IOException("Output directory could not be resolved: " + outputPath);
        }

        Path temporaryPath = Files.createTempFile(outputDirectory, TEMP_FILE_PREFIX, TEMP_FILE_SUFFIX);

        try {
            writeWorkbook(workbook, temporaryPath);

            replaceOutputFile(temporaryPath, absoluteOutputPath);

        } catch (IOException | RuntimeException | Error e) {
            deleteTemporaryFileSuppressingFailure(temporaryPath, e);

            throw e;
        }
    }

    private static void writeWorkbook(XSSFWorkbook workbook, Path path) throws IOException {

        try (OutputStream output = Files.newOutputStream(path)) {

            workbook.write(output);
        }
    }

    private static void replaceOutputFile(Path temporaryPath, Path outputPath) throws IOException {

        try {
            Files.move(temporaryPath, outputPath, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);

        } catch (AtomicMoveNotSupportedException e) {
            Files.move(temporaryPath, outputPath, StandardCopyOption.REPLACE_EXISTING);
        }
    }

    private static void deleteTemporaryFileSuppressingFailure(Path temporaryPath, Throwable originalFailure) {

        try {
            Files.deleteIfExists(temporaryPath);

        } catch (IOException cleanupFailure) {
            originalFailure.addSuppressed(cleanupFailure);
        }
    }
}