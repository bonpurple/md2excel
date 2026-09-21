package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

import md2excel.render.MarkdownRenderer.LineInfo;
import md2excel.render.MarkdownRenderer.LineKind;

public class MarkdownLineParserTest {

    @Test
    public void pipeLineIsNormalWhenTableParsingIsDisabled() {
        RenderState state = new RenderState(40, 1, 1);

        LineInfo line = MarkdownLineParser.parse("price | note", state, false);

        assertEquals(LineKind.NORMAL, line.kind);
        assertEquals("price | note", line.paragraphText);
    }

    @Test
    public void pipeLineIsTableWhenTableParsingIsEnabled() {
        RenderState state = new RenderState(40, 1, 1);

        LineInfo line = MarkdownLineParser.parse("h1 | h2", state, true);

        assertEquals(LineKind.TABLE_ROW, line.kind);
    }

    @Test
    public void nestedQuoteHeadingKeepsHeadingKind() {
        RenderState state = new RenderState(40, 1, 1);

        LineInfo line = MarkdownLineParser.parse("> > ### nested heading", state, false);

        assertEquals(LineKind.BLOCK_QUOTE, line.kind);
        assertEquals(2, line.getQuoteDepth());

        LineInfo content = line.getInnermostQuotedContent();

        assertEquals(LineKind.HEADING, content.kind);
        assertEquals(3, content.headingLevel);
        assertEquals("nested heading", content.headingText);
    }

    @Test
    public void nestedQuoteTableKeepsTableKind() {
        RenderState state = new RenderState(40, 1, 1);

        LineInfo line = MarkdownLineParser.parse("> > | a | b |", state, true);

        assertEquals(2, line.getQuoteDepth());
        assertTrue(line.isTableLike());

        assertEquals(LineKind.TABLE_ROW, line.getInnermostQuotedContent().kind);
    }

    @Test
    public void codeBlockContentIsNotParsedAsMarkdown() {
        RenderState state = new RenderState(40, 1, 1);

        state.codeBlock().open('`', 3, 0, false, -1);

        LineInfo line = MarkdownLineParser.parse("**not bold**", state, true);

        assertEquals(LineKind.CODE_LINE, line.kind);
        assertFalse(line.isTableLike());
    }

    @Test
    public void matchingClosingFenceIsRecognized() {
        RenderState state = new RenderState(40, 1, 1);

        state.codeBlock().open('`', 3, 0, false, -1);

        LineInfo line = MarkdownLineParser.parse("````", state, false);

        assertEquals(LineKind.CODE_FENCE, line.kind);
    }

    @Test
    public void differentFenceMarkerRemainsCodeContent() {
        RenderState state = new RenderState(40, 1, 1);

        state.codeBlock().open('`', 3, 0, false, -1);

        LineInfo line = MarkdownLineParser.parse("~~~", state, false);

        assertEquals(LineKind.CODE_LINE, line.kind);
    }
}