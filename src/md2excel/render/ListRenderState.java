package md2excel.render;

import java.util.ArrayList;
import java.util.List;

import md2excel.markdown.ListStackUtil;

/**
 * リスト階層とリスト継続状態を管理する。
 */
final class ListRenderState {

    private final List<ListStackUtil.ListLevel> levels = new ArrayList<ListStackUtil.ListLevel>();

    private boolean inListBlock;
    private boolean inNestedNumberBlock;
    private boolean bulletDetailActive;

    int updateDepth(int indent, boolean ordered) {
        return ListStackUtil.updateListDepth(levels, indent, ordered);
    }

    boolean hasLevels() {
        return !levels.isEmpty();
    }

    int getDepthForIndent(int indent) {
        return ListStackUtil.getDepthForIndent(levels, indent);
    }

    int getParentDepthForChildParagraph() {
        return ListStackUtil.getParentListDepthForChildParagraph(levels);
    }

    boolean isInListBlock() {
        return inListBlock;
    }

    void afterWriteBulletItem() {
        inNestedNumberBlock = false;
        bulletDetailActive = true;
        inListBlock = true;
    }

    void afterWriteNumberedItem() {
        bulletDetailActive = false;
        inNestedNumberBlock = true;
        inListBlock = true;
    }

    void afterWriteNormalText(boolean isListNote, int indent) {

        if (isListNote) {
            inListBlock = false;
        }

        if (indent == 0) {
            bulletDetailActive = false;
        }
    }

    /**
     * 段落やブロックが切り替わったとき、 箇条書きの説明行としての連結を解除する。
     */
    void cutParagraphLinking() {
        bulletDetailActive = false;
    }

    void resetOnBlockBoundary() {
        bulletDetailActive = false;
    }

    /**
     * 現在のリストブロックから離脱する。
     *
     * 後続行のインデント列計算に使用するため、 リスト階層levelsは既存仕様どおり保持する。
     */
    void leaveBlockPreservingLevels() {
        inListBlock = false;
        inNestedNumberBlock = false;
        bulletDetailActive = false;
    }

    boolean shouldInsertAutoBlankBeforeChildList(int currentIndent, boolean previousContentIsList) {

        if (levels.isEmpty()) {
            return false;
        }

        int previousIndent = levels.get(levels.size() - 1).indent;

        if (currentIndent >= previousIndent) {
            return false;
        }

        return previousContentIsList || bulletDetailActive || inNestedNumberBlock;
    }
}