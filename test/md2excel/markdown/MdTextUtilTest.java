package md2excel.markdown;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class MdTextUtilTest {

    @Test
    public void leadingTabCountsAsFourColumns() {
        assertEquals(6, MdTextUtil.countLeadingSpacesOrTabs("\t  text"));
    }

    @Test
    public void leadingIndentColumnsCanRemoveWholeTab() {
        assertEquals("text", MdTextUtil.removeLeadingIndentColumns("\ttext", 4));
    }

    @Test
    public void partialTabRemovalLeavesSpaces() {
        assertEquals("  text", MdTextUtil.removeLeadingIndentColumns("\ttext", 2));
    }

    @Test
    public void tabsAreExpandedToFourSpaces() {
        assertEquals("    x    y", MdTextUtil.expandTabs("\tx\ty"));
    }

    @Test
    public void detectsHardBreakByTrailingSpaces() {
        assertTrue(MdTextUtil.hasHardLineBreakBySpaces("text  "));

        assertFalse(MdTextUtil.hasHardLineBreakBySpaces("text "));
    }

    @Test
    public void detectsHardBreakByBackslash() {
        assertTrue(MdTextUtil.hasHardLineBreakByBackslash("text\\"));

        assertTrue(MdTextUtil.hasHardLineBreakByBackslash("text\\ \t"));

        assertFalse(MdTextUtil.hasHardLineBreakByBackslash("text"));
    }

    @Test
    public void recognizesOpeningAndClosingCodeFences() {
        assertTrue(MdTextUtil.isOpeningCodeFenceLine("```text"));

        assertTrue(MdTextUtil.isClosingCodeFenceLine("````  ", '`', 3));

        assertFalse(MdTextUtil.isClosingCodeFenceLine("``` trailing", '`', 3));

        assertFalse(MdTextUtil.isClosingCodeFenceLine("~~~", '`', 3));
    }

    @Test
    public void recognizesHorizontalRules() {
        assertTrue(MdTextUtil.isHorizontalRuleLine("---"));
        assertTrue(MdTextUtil.isHorizontalRuleLine("* * *"));
        assertTrue(MdTextUtil.isHorizontalRuleLine("_\t_\t_"));

        assertFalse(MdTextUtil.isHorizontalRuleLine("--"));
        assertFalse(MdTextUtil.isHorizontalRuleLine("-*-"));
    }
}