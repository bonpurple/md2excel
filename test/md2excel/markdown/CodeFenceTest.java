package md2excel.markdown;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class CodeFenceTest {

    @Test
    public void parsesBacktickOpeningFence() {
        CodeFence fence = CodeFence.parseOpening("```text");

        assertNotNull(fence);
        assertEquals('`', fence.getMarker());
        assertEquals(3, fence.getLength());
    }

    @Test
    public void parsesTildeOpeningFence() {
        CodeFence fence = CodeFence.parseOpening("~~~~java");

        assertNotNull(fence);
        assertEquals('~', fence.getMarker());
        assertEquals(4, fence.getLength());
    }

    @Test
    public void rejectsShortOpeningFence() {
        assertNull(CodeFence.parseOpening("``"));

        assertNull(CodeFence.parseOpening("~~"));
    }

    @Test
    public void recognizesClosingFenceOfSameLength() {
        assertTrue(CodeFence.isClosingLine("```", '`', 3));
    }

    @Test
    public void recognizesLongerClosingFence() {
        assertTrue(CodeFence.isClosingLine("````  ", '`', 3));
    }

    @Test
    public void rejectsShorterClosingFence() {
        assertFalse(CodeFence.isClosingLine("```", '`', 4));
    }

    @Test
    public void rejectsDifferentClosingMarker() {
        assertFalse(CodeFence.isClosingLine("~~~", '`', 3));
    }

    @Test
    public void rejectsTrailingNonWhitespace() {
        assertFalse(CodeFence.isClosingLine("``` trailing", '`', 3));
    }
}