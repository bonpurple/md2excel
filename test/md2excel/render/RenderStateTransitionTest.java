package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class RenderStateTransitionTest {

    @Test
    public void markdownBlankRemembersRowAndPreviousTable() {
        RenderState state = createState();

        state.afterWriteQuotedTableRow();

        int blankRow = state.allocateNextRowIndex();
        state.afterWriteMarkdownBlank(blankRow);

        assertEquals(blankRow, state.getLastBlankRowIndex());
        assertTrue(state.hasReusableMarkdownBlankForParagraph());
        assertTrue(state.isLastBlankAfterTable());
        assertFalse(state.isLastLineTable());

        // Markdown空行そのものは、直前が引用かどうかを変更しない。
        assertTrue(state.wasLastBlockQuote());
    }

    @Test
    public void quotedHeadingDoesNotStartNormalHeadingParagraphBlock() {
        RenderState state = createState();

        state.afterWriteQuotedHeading();

        assertTrue(state.wasLastBlockQuote());
        assertFalse(state.isInHeadingParagraphBlock());

        state.afterWriteHeading();

        assertFalse(state.wasLastBlockQuote());
        assertTrue(state.isInHeadingParagraphBlock());
    }

    @Test
    public void quotedCodeAndRuleKeepTheirOrigin() {
        RenderState state = createState();

        state.afterWriteQuotedCodeLine();
        assertTrue(state.wasLastBlockQuote());

        state.afterWriteCodeLine();
        assertFalse(state.wasLastBlockQuote());

        state.afterWriteQuotedHorizontalRule();
        assertTrue(state.wasLastBlockQuote());

        state.afterWriteHorizontalRule();
        assertFalse(state.wasLastBlockQuote());
    }

    @Test
    public void quotedTableTransitionsKeepTheirOrigin() {
        RenderState state = createState();

        state.afterWriteQuotedTableRow();
        assertTrue(state.wasLastBlockQuote());
        assertTrue(state.isLastLineTable());

        state.afterSkipQuotedTableSeparatorLine();
        assertTrue(state.wasLastBlockQuote());
        assertTrue(state.isLastLineTable());

        state.afterWriteTableRow();
        assertFalse(state.wasLastBlockQuote());
        assertTrue(state.isLastLineTable());
    }

    @Test
    public void normalListNoteStillLeavesListBlock() {
        RenderState state = createState();

        state.updateListDepth(0, false);
        state.afterWriteBulletItem();
        assertTrue(state.isInListBlock());

        state.afterWriteNormalText(0, true);
        assertFalse(state.isInListBlock());
    }

    @Test
    public void quotedTableRowRecordsCompleteInfoAndCanUpdateStyleRole() {
        RenderState state = createState();

        state.recordBlockQuoteTableRow(1, 2, 2, 3, 2, TableRowStyleRole.HEADER);

        assertTrue(state.wasLastBlockQuote());
        assertTrue(state.hasRenderableBlockQuote());
        assertEquals(1, state.getBlockQuoteFirstRow());
        assertEquals(1, state.getBlockQuoteLastRow());
        assertEquals(1, state.getBlockQuoteStartCol());

        QuoteRowInfo info = state.getBlockQuoteRowInfo(1);

        assertEquals(QuoteRowKind.TABLE, info.getKind());
        assertEquals(2, info.getDepth());
        assertEquals(2, info.getContentCol());
        assertEquals(TableRowStyleRole.HEADER, info.getTableRowStyleRole());
        assertFalse(info.isTableContentColumn(1));
        assertTrue(info.isTableContentColumn(2));
        assertTrue(info.isTableContentColumn(3));
        assertFalse(info.isTableContentColumn(4));

        state.updateBlockQuoteTableRowStyleRole(1, TableRowStyleRole.BODY_WITHOUT_BOTTOM_BORDER);

        QuoteRowInfo updated = state.getBlockQuoteRowInfo(1);
        assertEquals(TableRowStyleRole.BODY_WITHOUT_BOTTOM_BORDER, updated.getTableRowStyleRole());
        assertTrue(updated.isTableContentColumn(2));
        assertTrue(updated.isTableContentColumn(3));
    }

    @Test
    public void markdownBlankPreservesQuotedOriginUntilNormalTextIsWritten() {
        RenderState state = createState();

        int quoteRow = state.allocateNextRowIndex();
        state.afterWriteNormalText(0, false);
        state.recordBlockQuoteRow(quoteRow, 2, 2, QuoteRowKind.NORMAL, 1);

        assertTrue(state.wasLastBlockQuote());
        assertTrue(state.hasRenderableBlockQuote());

        int blankRow = state.allocateNextRowIndex();
        state.afterWriteMarkdownBlank(blankRow);

        assertTrue(state.wasLastBlockQuote());
        assertTrue(state.hasRenderableBlockQuote());
        assertTrue(state.hasReusableMarkdownBlankForParagraph());
        assertEquals(blankRow, state.getLastBlankRowIndex());

        state.allocateNextRowIndex();
        state.afterWriteNormalText(0, false);

        assertFalse(state.wasLastBlockQuote());
        assertFalse(state.hasReusableMarkdownBlankForParagraph());
    }

    @Test
    public void closingTableThenWritingMarkdownBlankRetainsTableOrigin() {
        RenderState state = createState();

        int headerRow = state.allocateNextRowIndex();
        state.table().begin(headerRow, 1, 2, 0);
        state.afterWriteTableRow();

        assertTrue(state.isLastLineTable());

        state.afterCloseTable();

        assertFalse(state.isLastLineTable());
        assertFalse(state.table().isOpen());

        int blankRow = state.allocateNextRowIndex();
        state.afterWriteMarkdownBlank(blankRow);

        assertTrue(state.isLastBlankAfterTable());
        assertTrue(state.hasReusableMarkdownBlankForParagraph());

        state.allocateNextRowIndex();
        state.afterWriteNormalText(0, false);

        assertFalse(state.isLastBlankAfterTable());
        assertFalse(state.hasReusableMarkdownBlankForParagraph());
    }

    private static RenderState createState() {
        return new RenderState(new SheetColumnLayout(40, 1, 1), 1);
    }
}