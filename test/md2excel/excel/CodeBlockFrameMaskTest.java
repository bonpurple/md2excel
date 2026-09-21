package md2excel.excel;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class CodeBlockFrameMaskTest {

    @Test
    public void combinedMaskContainsSelectedSides() {
        int mask = CodeBlockFrameMask.TOP | CodeBlockFrameMask.LEFT;

        assertTrue(CodeBlockFrameMask.contains(mask, CodeBlockFrameMask.TOP));

        assertTrue(CodeBlockFrameMask.contains(mask, CodeBlockFrameMask.LEFT));

        assertFalse(CodeBlockFrameMask.contains(mask, CodeBlockFrameMask.BOTTOM));

        assertFalse(CodeBlockFrameMask.contains(mask, CodeBlockFrameMask.RIGHT));
    }

    @Test
    public void allMaskContainsEverySide() {
        assertTrue(CodeBlockFrameMask.contains(CodeBlockFrameMask.ALL, CodeBlockFrameMask.TOP));

        assertTrue(CodeBlockFrameMask.contains(CodeBlockFrameMask.ALL, CodeBlockFrameMask.BOTTOM));

        assertTrue(CodeBlockFrameMask.contains(CodeBlockFrameMask.ALL, CodeBlockFrameMask.LEFT));

        assertTrue(CodeBlockFrameMask.contains(CodeBlockFrameMask.ALL, CodeBlockFrameMask.RIGHT));
    }

    @Test
    public void validatesMaskRange() {
        assertTrue(CodeBlockFrameMask.isValid(CodeBlockFrameMask.NONE));

        assertTrue(CodeBlockFrameMask.isValid(CodeBlockFrameMask.ALL));

        assertFalse(CodeBlockFrameMask.isValid(-1));

        assertFalse(CodeBlockFrameMask.isValid(CodeBlockFrameMask.COMBINATION_COUNT));
    }
}