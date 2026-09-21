package md2excel.render;

/**
 * 描画列範囲と、次に生成する行番号を管理する。
 */
final class RenderPositionState {

    private final SheetColumnLayout columnLayout;
    private final int startRowIndex;

    private int nextRowIndex;

    RenderPositionState(SheetColumnLayout columnLayout, int startRowIndex) {

        if (columnLayout == null) {
            throw new IllegalArgumentException("columnLayout must not be null");
        }

        this.columnLayout = columnLayout;
        this.startRowIndex = Math.max(0, startRowIndex);

        this.nextRowIndex = this.startRowIndex;
    }

    int getRenderEndColExclusive() {
        return columnLayout.getRenderEndColExclusive();
    }

    int getRenderLastColIndex() {
        return columnLayout.getRenderLastColIndex();
    }

    int getStartColIndex() {
        return columnLayout.getRenderStartColIndex();
    }

    int getNextRowIndex() {
        return nextRowIndex;
    }

    int allocateNextRowIndex() {
        return nextRowIndex++;
    }

    int getPreviousRowIndex() {
        return nextRowIndex - 1;
    }

    boolean hasWrittenRows() {
        return nextRowIndex > startRowIndex;
    }
}