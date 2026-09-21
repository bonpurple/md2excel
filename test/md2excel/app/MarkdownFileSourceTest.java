package md2excel.app;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

public class MarkdownFileSourceTest {

    @Rule
    public TemporaryFolder temporaryFolder = new TemporaryFolder();

    @Test
    public void readsUtf8Lines() throws Exception {
        Path path = temporaryFolder.newFile("input.md").toPath();

        Files.write(path, "日本語\nsecond".getBytes(StandardCharsets.UTF_8));

        try (MarkdownFileSource source = MarkdownFileSource.open(path)) {

            Iterator<String> iterator = source.iterator();

            assertTrue(iterator.hasNext());
            assertEquals("日本語", iterator.next());

            assertTrue(iterator.hasNext());
            assertEquals("second", iterator.next());

            assertFalse(iterator.hasNext());
        }
    }

    @Test
    public void iteratorCanOnlyBeCreatedOnce() throws Exception {

        Path path = temporaryFolder.newFile("input.md").toPath();

        Files.write(path, "text".getBytes(StandardCharsets.UTF_8));

        try (MarkdownFileSource source = MarkdownFileSource.open(path)) {

            source.iterator();

            try {
                source.iterator();
                fail("IllegalStateException was expected");

            } catch (IllegalStateException expected) {
                // expected
            }
        }
    }
}