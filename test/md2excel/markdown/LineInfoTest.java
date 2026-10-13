package md2excel.markdown;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class LineInfoTest {

    @Test
    public void simpleLineKeepsPreviousDefaultGetterValues() {
        LineInfo line = LineInfo.codeLine("code", "code", 0);

        assertEquals(LineKind.CODE_LINE, line.getKind());
        assertEquals(-1, line.getHeadingLevel());
        assertNull(line.getHeadingText());
        assertFalse(line.endsWithHardBreak());
        assertNull(line.getParagraphText());
        assertNull(line.getListMarkerText());
        assertNull(line.getListContentText());
    }

    @Test
    public void headingHasOnlyHeadingDetail() {
        LineInfo line = LineInfo.heading("# title", "# title", 0, 1, "title", true);

        assertEquals(1, line.getHeadingLevel());
        assertEquals("title", line.getHeadingText());
        assertTrue(line.endsWithHardBreak());
        assertNull(line.getParagraphText());
        assertNull(line.getListMarkerText());
        assertNull(line.getListContentText());
    }

    @Test
    public void normalHasOnlyParagraphDetail() {
        LineInfo line = LineInfo.normal("body", "body", 0, "body", true);

        assertEquals(-1, line.getHeadingLevel());
        assertNull(line.getHeadingText());
        assertEquals("body", line.getParagraphText());
        assertTrue(line.endsWithHardBreak());
        assertNull(line.getListMarkerText());
        assertNull(line.getListContentText());
    }

    @Test
    public void listKeepsContentAsParagraphText() {
        LineInfo line = LineInfo.numberItem("12. item", "12. item", 0, "12. ", "item", true);

        assertEquals(LineKind.NUMBER_ITEM, line.getKind());
        assertEquals(-1, line.getHeadingLevel());
        assertNull(line.getHeadingText());
        assertEquals("item", line.getParagraphText());
        assertEquals("12. ", line.getListMarkerText());
        assertEquals("item", line.getListContentText());
        assertTrue(line.endsWithHardBreak());
    }

    @Test
    public void quoteCopiesContentDataWithoutChangingRawLine() {
        LineInfo content = LineInfo.bulletItem("  - item", "- item", 2, "・ ", "item", true);

        LineInfo quoted = LineInfo.quoted(">   - item", ">   - item", 0, 1, content);

        assertTrue(quoted.isQuoted());
        assertEquals(1, quoted.getQuoteDepth());
        assertEquals(">   - item", quoted.getRaw());
        assertEquals("  - item", quoted.getContentRaw());
        assertEquals(2, quoted.getContentIndent());
        assertEquals(LineKind.BULLET_ITEM, quoted.getKind());
        assertEquals("item", quoted.getParagraphText());
        assertEquals("・ ", quoted.getListMarkerText());
        assertEquals("item", quoted.getListContentText());
        assertTrue(quoted.endsWithHardBreak());
    }
}