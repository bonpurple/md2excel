package md2excel.markdown;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import org.junit.Test;

public class NumberedListMarkerTest {

    @Test
    public void parsesDotMarker() {
        NumberedListMarker marker = NumberedListMarker.parse("12. text");

        assertEquals("12. ", marker.getMarkerText());

        assertEquals(4, marker.getContentStartIndex());
    }

    @Test
    public void parsesClosingParenthesisMarker() {
        NumberedListMarker marker = NumberedListMarker.parse("1) text");

        assertEquals("1) ", marker.getMarkerText());

        assertEquals(3, marker.getContentStartIndex());
    }

    @Test
    public void consumesAllWhitespaceAfterMarker() {
        NumberedListMarker marker = NumberedListMarker.parse("1. \t  text");

        assertEquals("1. ", marker.getMarkerText());

        assertEquals(6, marker.getContentStartIndex());
    }

    @Test
    public void acceptsMarkerWithOnlyTrailingWhitespace() {
        NumberedListMarker marker = NumberedListMarker.parse("1.   ");

        assertEquals(5, marker.getContentStartIndex());
    }

    @Test
    public void rejectsMarkerWithoutWhitespace() {
        assertNull(NumberedListMarker.parse("1.text"));
    }

    @Test
    public void rejectsNonNumericMarker() {
        assertNull(NumberedListMarker.parse("a. text"));
    }

    @Test
    public void rejectsUnsupportedPunctuation() {
        assertNull(NumberedListMarker.parse("1: text"));
    }
}