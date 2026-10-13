package md2excel.render;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class TableParsingProbeTest {

    @Test
    public void startsTableWhenNextLineIsMatchingSeparator() {
        RenderState state = createState();

        assertTrue(TableParsingProbe.shouldEnableTableParsing("h1 | h2", "--- | ---", state));
    }

    @Test
    public void doesNotStartTableWhenQuoteDepthsDiffer() {
        RenderState state = createState();

        assertFalse(TableParsingProbe.shouldEnableTableParsing("> h1 | h2", "> > --- | ---", state));
    }

    @Test
    public void startsQuotedTableWhenQuoteDepthsMatch() {
        RenderState state = createState();

        assertTrue(TableParsingProbe.shouldEnableTableParsing("> h1 | h2", "> --- | ---", state));
    }

    @Test
    public void doesNotStartTableWithoutNextLine() {
        RenderState state = createState();

        assertFalse(TableParsingProbe.shouldEnableTableParsing("h1 | h2", null, state));
    }

    @Test
    public void continuingTableDoesNotRequireNextLine() {
        RenderState state = createState();

        state.table().begin(1, 2, 3, 1);
        state.afterWriteQuotedTableRow();

        assertTrue(TableParsingProbe.shouldEnableTableParsing("> body | value", null, state));
    }

    @Test
    public void tableAtDifferentQuoteDepthDoesNotContinue() {
        RenderState state = createState();

        state.table().begin(1, 2, 3, 1);
        state.afterWriteQuotedTableRow();

        assertFalse(TableParsingProbe.shouldEnableTableParsing("> > body | value", null, state));
    }

    private static RenderState createState() {
        return new RenderState(new SheetColumnLayout(40, 1, 1), 1);
    }
}