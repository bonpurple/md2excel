package md2excel.render;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class MarkdownTableTest {

    @Test
    public void singlePipeLineAloneDoesNotStartTable() {
        assertFalse(MarkdownTable.isTableStart("price | note", "ordinary paragraph"));
    }

    @Test
    public void tableWithoutOuterPipesIsRecognized() {
        assertTrue(MarkdownTable.isTableStart("h1 | h2", "--- | ---"));
    }

    @Test
    public void tableWithOuterPipesIsRecognized() {
        assertTrue(MarkdownTable.isTableStart("| h1 | h2 |", "| --- | --- |"));
    }

    @Test
    public void alignmentMarkersAreAccepted() {
        assertTrue(MarkdownTable.isTableSeparatorLine("| :--- | ---: |"));
    }

    @Test
    public void mismatchedColumnCountDoesNotStartTable() {
        assertFalse(MarkdownTable.isTableStart("h1 | h2", "--- | --- | ---"));
    }

    @Test
    public void invalidSeparatorDoesNotStartTable() {
        assertFalse(MarkdownTable.isTableStart("h1 | h2", "abc | ---"));
    }
}