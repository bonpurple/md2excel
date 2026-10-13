package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.List;

import org.junit.Test;

public class MarkdownInlineTest {

    @Test
    public void parsesPlainText() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("plain text");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "plain text", false, false, false);
    }

    @Test
    public void parsesStrongEmphasis() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("**bold**");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "bold", true, false, false);
    }

    @Test
    public void parsesEmphasis() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("*italic*");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "italic", false, true, false);
    }

    @Test
    public void parsesNestedStrongAndEmphasis() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("***bold italic***");

        assertEquals(1, segments.size());

        MdSegment segment = segments.get(0);

        assertEquals("bold italic", segment.text);
        assertTrue(segment.inBold);
        assertTrue(segment.inItalic);
        assertFalse(segment.inCode);
    }

    @Test
    public void parsesStrongContainingEmphasis() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("**bold *italic* bold**");

        assertEquals(3, segments.size());

        assertSegment(segments.get(0), "bold ", true, false, false);

        assertSegment(segments.get(1), "italic", true, true, false);

        assertSegment(segments.get(2), " bold", true, false, false);
    }

    @Test
    public void escapedAsteriskIsPlainText() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("\\*not italic\\*");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "*not italic*", false, false, false);
    }

    @Test
    public void escapedAsteriskInsideStrongIsPlainCharacter() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("**a\\*b**");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "a*b", true, false, false);
    }

    @Test
    public void inlineCodeDoesNotParseEmphasis() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("`*not italic*`");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "*not italic*", false, false, true);
    }

    @Test
    public void multipleBackticksAllowSingleBacktickInsideCode() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("``a`b``");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "a`b", false, false, true);
    }

    @Test
    public void unmatchedBackticksRemainPlainText() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("`unclosed");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "`unclosed", false, false, false);
    }

    @Test
    public void brOutsideCodeCreatesDisplayLine() {
        List<List<MdSegment>> lines = MarkdownInlineParser.parseParagraphToDisplayLines("foo<br>**bar**");

        assertEquals(2, lines.size());

        assertEquals(1, lines.get(0).size());
        assertSegment(lines.get(0).get(0), "foo", false, false, false);

        assertEquals(1, lines.get(1).size());
        assertSegment(lines.get(1).get(0), "bar", true, false, false);
    }

    @Test
    public void brInsideInlineCodeRemainsLiteral() {
        List<List<MdSegment>> lines = MarkdownInlineParser.parseParagraphToDisplayLines("`aa<br>bb` and **cc**");

        assertEquals(1, lines.size());

        assertEquals(3, lines.get(0).size());

        assertSegment(lines.get(0).get(0), "aa<br>bb", false, false, true);

        assertSegment(lines.get(0).get(1), " and ", false, false, false);

        assertSegment(lines.get(0).get(2), "cc", true, false, false);
    }

    @Test
    public void softBreakBecomesSingleSpace() {
        String paragraph = "abc" + ParagraphBuffer.SOFT_BREAK_TOKEN + "def";

        List<List<MdSegment>> lines = MarkdownInlineParser.parseParagraphToDisplayLines(paragraph);

        assertEquals(1, lines.size());
        assertEquals(1, lines.get(0).size());

        assertSegment(lines.get(0).get(0), "abc def", false, false, false);
    }

    @Test
    public void hardBreakCreatesNextDisplayLine() {
        String paragraph = "abc" + ParagraphBuffer.HARD_BREAK_TOKEN + "**def**";

        List<List<MdSegment>> lines = MarkdownInlineParser.parseParagraphToDisplayLines(paragraph);

        assertEquals(2, lines.size());

        assertSegment(lines.get(0).get(0), "abc", false, false, false);

        assertSegment(lines.get(1).get(0), "def", true, false, false);
    }

    @Test
    public void softBreakInsideInlineCodeBecomesSpace() {
        String paragraph = "`aa" + ParagraphBuffer.SOFT_BREAK_TOKEN + "bb`";

        List<List<MdSegment>> lines = MarkdownInlineParser.parseParagraphToDisplayLines(paragraph);

        assertEquals(1, lines.size());

        assertSegment(lines.get(0).get(0), "aa bb", false, false, true);
    }

    @Test
    public void hardBreakInsideInlineCodeBecomesSpace() {
        String paragraph = "`aa" + ParagraphBuffer.HARD_BREAK_TOKEN + "bb`";

        List<List<MdSegment>> lines = MarkdownInlineParser.parseParagraphToDisplayLines(paragraph);

        assertEquals(1, lines.size());

        assertSegment(lines.get(0).get(0), "aa bb", false, false, true);
    }

    @Test
    public void underscoreInsideWordDoesNotCreateEmphasis() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("foo_bar_baz");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "foo_bar_baz", false, false, false);
    }

    @Test
    public void unmatchedEmphasisDelimiterRemainsText() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("*unclosed");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "*unclosed", false, false, false);
    }

    @Test
    public void ruleOfThreeDoesNotCreateInvalidEmphasis() {
        List<MdSegment> segments = MarkdownInlineParser.parseParagraphToSingleLineSegments("a**b*c");

        assertEquals("a**b*c", toPlainText(segments));
    }

    @Test
    public void tokenizationEmphasisAndDisplaySplitWorkTogether() {
        List<List<MdSegment>> lines = MarkdownInlineParser.parseParagraphToDisplayLines("*a* and `b`<br>**c**");

        assertEquals(2, lines.size());
        assertEquals(3, lines.get(0).size());

        assertSegment(lines.get(0).get(0), "a", false, true, false);
        assertSegment(lines.get(0).get(1), " and ", false, false, false);
        assertSegment(lines.get(0).get(2), "b", false, false, true);

        assertEquals(1, lines.get(1).size());
        assertSegment(lines.get(1).get(0), "c", true, false, false);
    }

    private static void assertSegment(MdSegment segment, String text, boolean bold, boolean italic, boolean code) {

        assertEquals(text, segment.text);
        assertEquals(bold, segment.inBold);
        assertEquals(italic, segment.inItalic);
        assertEquals(code, segment.inCode);
    }

    private static String toPlainText(List<MdSegment> segments) {

        StringBuilder out = new StringBuilder();

        for (MdSegment segment : segments) {
            out.append(segment.text);
        }

        return out.toString();
    }
}