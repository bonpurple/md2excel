package md2excel.markdown;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

public class ListStackUtilTest {

    @Test
    public void updatesNestedListDepth() {
        List<ListStackUtil.ListLevel> stack = new ArrayList<ListStackUtil.ListLevel>();

        assertEquals(0, ListStackUtil.updateListDepth(stack, 0, false));

        assertEquals(1, ListStackUtil.updateListDepth(stack, 2, false));

        assertEquals(2, ListStackUtil.updateListDepth(stack, 4, true));

        assertEquals(3, stack.size());
        assertTrue(stack.get(2).ordered);
    }

    @Test
    public void returningToRootRemovesNestedLevels() {
        List<ListStackUtil.ListLevel> stack = new ArrayList<ListStackUtil.ListLevel>();

        ListStackUtil.updateListDepth(stack, 0, false);
        ListStackUtil.updateListDepth(stack, 2, false);
        ListStackUtil.updateListDepth(stack, 4, true);

        int depth = ListStackUtil.updateListDepth(stack, 0, false);

        assertEquals(0, depth);
        assertEquals(1, stack.size());
        assertFalse(stack.get(0).ordered);
    }

    @Test
    public void calculatesDepthForIntermediateIndent() {
        List<ListStackUtil.ListLevel> stack = new ArrayList<ListStackUtil.ListLevel>();

        stack.add(new ListStackUtil.ListLevel(0, false));
        stack.add(new ListStackUtil.ListLevel(2, false));

        assertEquals(0, ListStackUtil.getDepthForIndent(stack, 0));

        assertEquals(1, ListStackUtil.getDepthForIndent(stack, 1));

        assertEquals(2, ListStackUtil.getDepthForIndent(stack, 3));
    }

    @Test
    public void orderedLevelIsUsedAsParentForChildParagraph() {
        List<ListStackUtil.ListLevel> stack = new ArrayList<ListStackUtil.ListLevel>();

        stack.add(new ListStackUtil.ListLevel(0, false));
        stack.add(new ListStackUtil.ListLevel(2, true));
        stack.add(new ListStackUtil.ListLevel(4, false));

        assertEquals(1, ListStackUtil.getParentListDepthForChildParagraph(stack));
    }
}