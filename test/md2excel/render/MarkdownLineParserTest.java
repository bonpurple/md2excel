package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

import md2excel.markdown.LineInfo;
import md2excel.markdown.LineKind;

public class MarkdownLineParserTest {

    @Test
    public void pipeLineIsNormalWhenTableParsingIsDisabled() {
        RenderState state = createState();

        LineInfo line = MarkdownLineParser.parse("price | note", state, false);

        assertEquals(LineKind.NORMAL, line.getKind());

        assertEquals("price | note", line.getParagraphText());
    }

    @Test
    public void pipeLineIsTableWhenTableParsingIsEnabled() {
        RenderState state = createState();

        LineInfo line = MarkdownLineParser.parse("h1 | h2", state, true);

        assertEquals(LineKind.TABLE_ROW, line.getKind());
    }

    @Test
    public void nestedQuoteHeadingKeepsHeadingKind() {
        RenderState state = createState();

        LineInfo line = MarkdownLineParser.parse("> > ### nested heading", state, false);

        assertTrue(line.isQuoted());

        assertEquals(2, line.getQuoteDepth());

        assertEquals(LineKind.HEADING, line.getKind());

        assertEquals(3, line.getHeadingLevel());

        assertEquals("nested heading", line.getHeadingText());

        assertEquals("> > ### nested heading", line.getRaw());

        assertEquals("### nested heading", line.getContentRaw());
    }

    @Test
    public void nestedQuoteTableKeepsTableKind() {
        RenderState state = createState();

        LineInfo line = MarkdownLineParser.parse("> > | a | b |", state, true);

        assertTrue(line.isQuoted());
        assertEquals(2, line.getQuoteDepth());
        assertTrue(line.isTableLike());

        assertEquals(LineKind.TABLE_ROW, line.getKind());

        assertEquals("| a | b |", line.getContentRaw());
    }

    @Test
    public void codeBlockContentIsNotParsedAsMarkdown() {
        RenderState state = createState();

        state.codeBlock().open('`', 3, 0, false, -1, 0);

        LineInfo line = MarkdownLineParser.parse("**not bold**", state, true);

        assertEquals(LineKind.CODE_LINE, line.getKind());
        assertFalse(line.isTableLike());
    }

    @Test
    public void matchingClosingFenceIsRecognized() {
        RenderState state = createState();

        state.codeBlock().open('`', 3, 0, false, -1, 0);

        LineInfo line = MarkdownLineParser.parse("````", state, false);

        assertEquals(LineKind.CODE_FENCE, line.getKind());
    }

    @Test
    public void differentFenceMarkerRemainsCodeContent() {
        RenderState state = createState();

        state.codeBlock().open('`', 3, 0, false, -1, 0);

        LineInfo line = MarkdownLineParser.parse("~~~", state, false);

        assertEquals(LineKind.CODE_LINE, line.getKind());
    }

    @Test
    public void nestedQuotedCodeBlockRemovesAllQuoteMarkers() {
        RenderState state = createState();

        state.codeBlock().open('`', 3, 0, true, 2, 2);

        LineInfo line = MarkdownLineParser.parse("> > code", state, false);

        assertTrue(line.isQuoted());

        assertEquals(2, line.getQuoteDepth());

        assertEquals(LineKind.CODE_LINE, line.getKind());

        assertEquals("code", line.getContentRaw());
    }

    @Test
    public void nestedQuotedClosingFenceIsRecognized() {
        RenderState state = createState();

        state.codeBlock().open('`', 3, 0, true, 2, 2);

        LineInfo line = MarkdownLineParser.parse("> > ```", state, false);

        assertEquals(2, line.getQuoteDepth());

        assertTrue(line.isQuoted());

        assertEquals(2, line.getQuoteDepth());

        assertEquals(LineKind.CODE_FENCE, line.getKind());

        assertEquals("```", line.getContentRaw());
    }

    @Test
    public void quotedLineKeepsHardBreakFromContent() {
        RenderState state = createState();

        LineInfo line = MarkdownLineParser.parse("> text\\", state, false);

        assertTrue(line.isQuoted());

        assertEquals(1, line.getQuoteDepth());

        assertEquals(LineKind.NORMAL, line.getKind());

        assertTrue(line.endsWithHardBreak());

        assertEquals("text\\", line.getContentRaw());

        assertEquals("text", line.getParagraphText());
    }

    @Test
    public void quotedCodeKeepsGreaterThanBeyondOpeningDepth() {
        RenderState state = createState();

        state.codeBlock().open('`', 3, 0, true, 2, 2);

        LineInfo line = MarkdownLineParser.parse("> > > literal greater-than", state, false);

        assertTrue(line.isQuoted());

        assertEquals(2, line.getQuoteDepth());

        assertEquals(LineKind.CODE_LINE, line.getKind());

        assertEquals("> literal greater-than", line.getContentRaw());
    }

    @Test
    public void nestedQuoteIsRepresentedWithoutRecursiveLineInfo() {
        RenderState state = createState();

        LineInfo line = MarkdownLineParser.parse("> > paragraph", state, false);

        assertTrue(line.isQuoted());
        assertEquals(2, line.getQuoteDepth());
        assertEquals(LineKind.NORMAL, line.getKind());

        assertEquals("> > paragraph", line.getRaw());

        assertEquals("paragraph", line.getContentRaw());

        assertEquals("paragraph", line.getParagraphText());
    }

    private static RenderState createState() {
        return new RenderState(new SheetColumnLayout(40, 1, 1), 1);
    }
}