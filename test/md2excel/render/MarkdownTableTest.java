package md2excel.render;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertEquals;
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

    // R1前の実装を実行して観察した境界。Markdown仕様への適合性は検証しない。
    @Test
    public void pipeRecognitionAndHeaderWidthDependOnBackslashParity() {
        String[] backslashes = { "\\", "\\\\", "\\\\\\", "\\\\\\\\" };
        boolean[] recognized = { false, true, false, true };
        for (int i = 0; i < backslashes.length; i++) {
            String label = "Backslashes before pipe: " + (i + 1);
            String line = "a" + backslashes[i] + "|b";
            assertEquals(label, recognized[i], MarkdownTable.isTableLine(line));
            assertEquals(label, recognized[i], MarkdownTable.isTableStart(line, "-|-"));
            assertEquals(label + ", one column", !recognized[i], MarkdownTable.isTableStart("|" + line + "|", "|-|"));
            assertEquals(label + ", two columns", recognized[i], MarkdownTable.isTableStart("|" + line + "|", "|-|-|"));
        }
    }

    @Test
    public void unescapedPipeInsideBackticksCountsAsHeaderDelimiter() {
        assertTrue(MarkdownTable.isTableLine("`a|b`"));
        assertTrue(MarkdownTable.isTableStart("`a|b`", "-|-"));
    }

    @Test
    public void outerPipesAndEmptyCellsDetermineHeaderWidth() {
        String[] headers = { "|a|b", "a|b|", "|a||c|", "|a|b\\|", "||b|", "|a||" };
        boolean[] twoColumns = { true, true, false, true, true, true };
        for (int i = 0; i < headers.length; i++) {
            assertEquals(headers[i] + ", two columns", twoColumns[i], MarkdownTable.isTableStart(headers[i], "|-|-|"));
            assertEquals(headers[i] + ", three columns", !twoColumns[i],
                    MarkdownTable.isTableStart(headers[i], "|-|-|-|"));
        }
    }

    @Test
    public void singleHyphenWithOptionalColonsIsAccepted() {
        for (String separator : new String[] { "|-|", "|:-|", "|-:|", "|:-:|" }) {
            assertTrue(separator, MarkdownTable.isTableSeparatorLine(separator));
        }
    }

    @Test
    public void colonsWithoutHyphenAreNotSeparators() {
        for (String separator : new String[] { "|:|", "|::|" }) {
            assertFalse(separator, MarkdownTable.isTableSeparatorLine(separator));
        }
    }

    @Test
    public void separatorIgnoresWhitespaceButNotNonBreakingSpace() {
        assertTrue(MarkdownTable.isTableSeparatorLine("| - \t-\u3000- |"));
        assertFalse(MarkdownTable.isTableSeparatorLine("| -\u00a0- |"));
    }

    @Test
    public void emptySeparatorCellsAreRejectedAtEveryPosition() {
        for (String separator : new String[] { "| |", "||-|", "|-||-|", "|-||" }) {
            assertFalse(separator, MarkdownTable.isTableSeparatorLine(separator));
        }
    }

    @Test
    public void separatorCannotBeUsedAsHeader() {
        assertFalse(MarkdownTable.isTableStart("|-|", "|-|"));
    }

    @Test(expected = NullPointerException.class)
    public void nullTableLineThrowsNullPointerException() {
        MarkdownTable.isTableLine(null);
    }

    @Test(expected = NullPointerException.class)
    public void nullHeaderThrowsNullPointerException() {
        MarkdownTable.isTableStart(null, "-|-|");
    }

    @Test(expected = NullPointerException.class)
    public void nullHeaderAndSeparatorThrowNullPointerException() {
        MarkdownTable.isTableStart(null, null);
    }

    @Test
    public void nullSeparatorIsRejected() {
        assertFalse(MarkdownTable.isTableSeparatorLine(null));
        assertFalse(MarkdownTable.isTableStart("a|b", null));
    }

    @Test
    public void emptyAndWhitespaceOnlyInputsAreRejected() {
        for (String input : new String[] { "", " \t " }) {
            String label = "Input length: " + input.length();
            assertFalse(label, MarkdownTable.isTableLine(input));
            assertFalse(label, MarkdownTable.isTableSeparatorLine(input));
            assertFalse(label + ", header", MarkdownTable.isTableStart(input, "-|-"));
            assertFalse(label + ", separator", MarkdownTable.isTableStart("a|b", input));
            assertFalse(label + ", both", MarkdownTable.isTableStart(input, input));
        }
    }
}
