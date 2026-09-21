package md2excel.app;

import java.io.IOException;
import java.io.UncheckedIOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import md2excel.config.Md2ExcelConfig;

/**
 * Markdown入力、Workbook描画、Excel出力を調整する。
 */
public final class MarkdownToExcelConverter {

    private final MarkdownWorkbookRenderer renderer;
    private final AtomicWorkbookWriter writer;

    public MarkdownToExcelConverter() {
        this(new MarkdownWorkbookRenderer(), new AtomicWorkbookWriter());
    }

    MarkdownToExcelConverter(MarkdownWorkbookRenderer renderer, AtomicWorkbookWriter writer) {

        if (renderer == null) {
            throw new IllegalArgumentException("renderer must not be null");
        }

        if (writer == null) {
            throw new IllegalArgumentException("writer must not be null");
        }

        this.renderer = renderer;
        this.writer = writer;
    }

    public void convert(Md2ExcelConfig config) throws IOException {

        if (config == null) {
            throw new IllegalArgumentException("config must not be null");
        }

        try (MarkdownFileSource source = MarkdownFileSource.open(config.getInputPath());
                XSSFWorkbook workbook = new XSSFWorkbook()) {

            renderer.render(source.iterator(), workbook, config.getFontSettings(), config.getSheetSettings());

            writer.write(workbook, config.getOutputPath());

        } catch (UncheckedIOException e) {
            /*
             * Files.lines()の遅延読み込み中に発生したIOExceptionを、 convert()のAPIに合わせてIOExceptionへ戻す。
             */
            throw e.getCause();
        }
    }
}