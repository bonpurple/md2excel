package md2excel.markdown;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class MdInlineCodeUtilTest {

    @Test
    public void countsBacktickRun() {
        assertEquals(3, MdInlineCodeUtil.countBackticks("```code", 0));

        assertEquals(2, MdInlineCodeUtil.countBackticks("a``b", 1));
    }

    @Test
    public void findsClosingRunWithSameLength() {
        String text = "``a`b``";

        assertEquals(5, MdInlineCodeUtil.findClosingBackticks(text, 2, 2));
    }

    @Test
    public void brOutsideCodeSpanIsReplaced() {
        assertEquals("`aa<br>bb`|cc", MdInlineCodeUtil.replaceBrOutsideCodeSpans("`aa<br>bb`<br>cc", "|"));
    }

    @Test
    public void brInsideMultipleBacktickCodeSpanIsPreserved() {
        assertEquals("``aa`<br>bb``|x", MdInlineCodeUtil.replaceBrOutsideCodeSpans("``aa`<br>bb``<br>x", "|"));
    }

    @Test
    public void escapedOpeningAngleBracketIsNotBr() {
        assertEquals("\\<br> x", MdInlineCodeUtil.replaceBrOutsideCodeSpans("\\<br> x", "|"));
    }

    @Test
    public void recognizesBrVariants() {
        assertEquals(4, MdInlineCodeUtil.matchBrTagLength("<br>", 0));

        assertEquals(5, MdInlineCodeUtil.matchBrTagLength("<BR/>", 0));

        assertEquals(6, MdInlineCodeUtil.matchBrTagLength("<br />", 0));

        assertEquals(0, MdInlineCodeUtil.matchBrTagLength("<break>", 0));
    }
}