package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.List;

import org.junit.Test;

public class MarkdownInlineTest {

    @Test
    public void parsesPlainText() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("plain text");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "plain text", false, false, false);
    }

    @Test
    public void parsesStrongEmphasis() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("**bold**");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "bold", true, false, false);
    }

    @Test
    public void parsesEmphasis() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("*italic*");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "italic", false, true, false);
    }

    @Test
    public void parsesNestedStrongAndEmphasis() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline
                .parseParagraphToSingleLineSegments("***bold italic***");

        assertEquals(1, segments.size());

        MarkdownInline.MdSegment segment = segments.get(0);

        assertEquals("bold italic", segment.text);
        assertTrue(segment.inBold);
        assertTrue(segment.inItalic);
        assertFalse(segment.inCode);
    }

    @Test
    public void parsesStrongContainingEmphasis() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline
                .parseParagraphToSingleLineSegments("**bold *italic* bold**");

        assertEquals(3, segments.size());

        assertSegment(segments.get(0), "bold ", true, false, false);

        assertSegment(segments.get(1), "italic", true, true, false);

        assertSegment(segments.get(2), " bold", true, false, false);
    }

    @Test
    public void escapedAsteriskIsPlainText() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("\\*not italic\\*");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "*not italic*", false, false, false);
    }

    @Test
    public void escapedAsteriskInsideStrongIsPlainCharacter() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("**a\\*b**");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "a*b", true, false, false);
    }

    @Test
    public void inlineCodeDoesNotParseEmphasis() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("`*not italic*`");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "*not italic*", false, false, true);
    }

    @Test
    public void multipleBackticksAllowSingleBacktickInsideCode() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("``a`b``");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "a`b", false, false, true);
    }

    @Test
    public void unmatchedBackticksRemainPlainText() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("`unclosed");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "`unclosed", false, false, false);
    }

    @Test
    public void brOutsideCodeCreatesDisplayLine() {
        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines("foo<br>**bar**");

        assertEquals(2, lines.size());

        assertEquals(1, lines.get(0).size());
        assertSegment(lines.get(0).get(0), "foo", false, false, false);

        assertEquals(1, lines.get(1).size());
        assertSegment(lines.get(1).get(0), "bar", true, false, false);
    }

    @Test
    public void brInsideInlineCodeRemainsLiteral() {
        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline
                .parseParagraphToDisplayLines("`aa<br>bb` and **cc**");

        assertEquals(1, lines.size());

        assertEquals(3, lines.get(0).size());

        assertSegment(lines.get(0).get(0), "aa<br>bb", false, false, true);

        assertSegment(lines.get(0).get(1), " and ", false, false, false);

        assertSegment(lines.get(0).get(2), "cc", true, false, false);
    }

    @Test
    public void softBreakBecomesSingleSpace() {
        String paragraph = "abc" + ParagraphBuffer.SOFT_BREAK_TOKEN + "def";

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(paragraph);

        assertEquals(1, lines.size());
        assertEquals(1, lines.get(0).size());

        assertSegment(lines.get(0).get(0), "abc def", false, false, false);
    }

    @Test
    public void hardBreakCreatesNextDisplayLine() {
        String paragraph = "abc" + ParagraphBuffer.HARD_BREAK_TOKEN + "**def**";

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(paragraph);

        assertEquals(2, lines.size());

        assertSegment(lines.get(0).get(0), "abc", false, false, false);

        assertSegment(lines.get(1).get(0), "def", true, false, false);
    }

    @Test
    public void softBreakInsideInlineCodeBecomesSpace() {
        String paragraph = "`aa" + ParagraphBuffer.SOFT_BREAK_TOKEN + "bb`";

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(paragraph);

        assertEquals(1, lines.size());

        assertSegment(lines.get(0).get(0), "aa bb", false, false, true);
    }

    @Test
    public void hardBreakInsideInlineCodeBecomesSpace() {
        String paragraph = "`aa" + ParagraphBuffer.HARD_BREAK_TOKEN + "bb`";

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(paragraph);

        assertEquals(1, lines.size());

        assertSegment(lines.get(0).get(0), "aa bb", false, false, true);
    }

    @Test
    public void underscoreInsideWordDoesNotCreateEmphasis() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("foo_bar_baz");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "foo_bar_baz", false, false, false);
    }

    @Test
    public void unmatchedEmphasisDelimiterRemainsText() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("*unclosed");

        assertEquals(1, segments.size());

        assertSegment(segments.get(0), "*unclosed", false, false, false);
    }

    @Test
    public void ruleOfThreeDoesNotCreateInvalidEmphasis() {
        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments("a**b*c");

        assertEquals("a**b*c", toPlainText(segments));
    }

    private static void assertSegment(MarkdownInline.MdSegment segment, String text, boolean bold, boolean italic,
            boolean code) {

        assertEquals(text, segment.text);
        assertEquals(bold, segment.inBold);
        assertEquals(italic, segment.inItalic);
        assertEquals(code, segment.inCode);
    }

    private static String toPlainText(List<MarkdownInline.MdSegment> segments) {

        StringBuilder out = new StringBuilder();

        for (MarkdownInline.MdSegment segment : segments) {
            out.append(segment.text);
        }

        return out.toString();
    }
}