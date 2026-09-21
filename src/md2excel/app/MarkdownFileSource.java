package md2excel.app;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.stream.Stream;

/**
 * UTF-8のMarkdownファイルから行を読み込む。
 */
final class MarkdownFileSource implements AutoCloseable {

    private final Stream<String> lines;
    private boolean iteratorCreated;

    private MarkdownFileSource(Stream<String> lines) {
        this.lines = lines;
    }

    static MarkdownFileSource open(Path inputPath) throws IOException {

        if (inputPath == null) {
            throw new IllegalArgumentException("inputPath must not be null");
        }

        return new MarkdownFileSource(Files.lines(inputPath, StandardCharsets.UTF_8));
    }

    Iterator<String> iterator() {
        if (iteratorCreated) {
            throw new IllegalStateException("iterator has already been created");
        }

        iteratorCreated = true;
        return lines.iterator();
    }

    @Override
    public void close() {
        lines.close();
    }
}