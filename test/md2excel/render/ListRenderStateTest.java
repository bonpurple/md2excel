package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class ListRenderStateTest {

    @Test
    public void leaveBlockPreservesListLevels() {
        ListRenderState state = new ListRenderState();

        assertEquals(0, state.updateDepth(0, false));

        assertEquals(1, state.updateDepth(2, true));

        state.afterWriteNumberedItem();

        assertTrue(state.isInListBlock());
        assertTrue(state.hasLevels());

        state.leaveBlockPreservingLevels();

        assertFalse(state.isInListBlock());

        // リストブロック状態は解除するが、
        // インデント計算用の階層は保持する。
        assertTrue(state.hasLevels());

        assertEquals(2, state.getDepthForIndent(3));
    }

    @Test
    public void leaveBlockClearsContinuationState() {
        ListRenderState state = new ListRenderState();

        state.updateDepth(0, false);

        state.updateDepth(2, true);

        state.afterWriteNumberedItem();

        state.leaveBlockPreservingLevels();

        /*
         * previousContentIsListがfalseの場合、 離脱後の内部継続状態だけを理由として 自動空行を要求しない。
         */
        assertFalse(state.shouldInsertAutoBlankBeforeChildList(0, false));
    }
}