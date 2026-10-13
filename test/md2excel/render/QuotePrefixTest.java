package md2excel.render;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class QuotePrefixTest {

    @Test
    public void unwrapsAllMarkersForNormalParsingAndTableProbe() {
        QuotePrefix prefix = QuotePrefix.unwrap("> > text", Integer.MAX_VALUE);

        assertEquals(2, prefix.depth);
        assertEquals("text", prefix.contentRaw);
    }

    @Test
    public void leavesMarkersBeyondQuotedCodeOpeningDepth() {
        QuotePrefix prefix = QuotePrefix.unwrap("> > > literal", 2);

        assertEquals(2, prefix.depth);
        assertEquals("> literal", prefix.contentRaw);
    }

    @Test
    public void leavesNonQuotedLineUnchanged() {
        QuotePrefix prefix = QuotePrefix.unwrap("  text", Integer.MAX_VALUE);

        assertEquals(0, prefix.depth);
        assertEquals("  text", prefix.contentRaw);
    }
}