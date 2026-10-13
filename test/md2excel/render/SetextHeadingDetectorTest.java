package md2excel.render;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class SetextHeadingDetectorTest {

    @Test
    public void detectsEqualsAndDashUnderlines() {
        assertEquals(1, detect("==="));
        assertEquals(2, detect("--- \t"));
    }

    @Test
    public void acceptsAtMostThreeLeadingSpaces() {
        assertEquals(2, detect("   ---"));
        assertEquals(0, detect("    ---"));
    }

    @Test
    public void rejectsLeadingTabAndOtherTrailingCharacters() {
        assertEquals(0, detect("\t---"));
        assertEquals(0, detect("--- text"));
    }

    @Test
    public void rejectsMissingOrIneligibleParagraph() {
        assertEquals(0, SetextHeadingDetector.detect("===", false, true, 0));
        assertEquals(0, SetextHeadingDetector.detect("===", true, false, 0));
        assertEquals(0, SetextHeadingDetector.detect("===", true, true, 4));
    }

    @Test
    public void rejectsEmptyUnderline() {
        assertEquals(0, detect(""));
        assertEquals(0, detect(null));
    }

    private static int detect(String raw) {
        return SetextHeadingDetector.detect(raw, true, true, 0);
    }
}