package md2excel.markdown;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class MdCharUtilTest {

    @Test
    public void recognizesAsciiPunctuation() {
        assertTrue(MdCharUtil.isAsciiPunctuation('*'));
        assertTrue(MdCharUtil.isAsciiPunctuation('_'));
        assertTrue(MdCharUtil.isAsciiPunctuation('`'));
        assertTrue(MdCharUtil.isAsciiPunctuation('\\'));
        assertTrue(MdCharUtil.isAsciiPunctuation('<'));
    }

    @Test
    public void rejectsLettersDigitsAndWhitespace() {
        assertFalse(MdCharUtil.isAsciiPunctuation('a'));
        assertFalse(MdCharUtil.isAsciiPunctuation('1'));
        assertFalse(MdCharUtil.isAsciiPunctuation(' '));
        assertFalse(MdCharUtil.isAsciiPunctuation('\t'));
    }

    @Test
    public void rejectsNonAsciiCharacters() {
        assertFalse(MdCharUtil.isAsciiPunctuation('あ'));
    }
}