package md2excel.render;

import java.util.HashMap;
import java.util.Map;

/**
 * 引用ブロックの描画範囲と、行ごとの引用情報を管理する。
 */
final class BlockQuoteState {

    private boolean open;

    private int firstRow = -1;
    private int lastRow = -1;
    private int startCol;

    /**
     * 直前に書いた内容が引用ブロックだったか。
     *
     * 引用ブロック終了後の自動空行判定にも使用するため、 clearTracking()ではリセットしない。
     */
    private boolean lastWasBlockQuote;

    private final Map<Integer, RenderState.QuoteRowInfo> rows = new HashMap<Integer, RenderState.QuoteRowInfo>();

    boolean isOpen() {
        return open;
    }

    boolean hasRenderableRows() {
        return open && firstRow >= 0 && lastRow >= 0;
    }

    int getFirstRow() {
        return firstRow;
    }

    int getLastRow() {
        return lastRow;
    }

    int getStartCol() {
        return startCol;
    }

    boolean wasLastBlockQuote() {
        return lastWasBlockQuote;
    }

    void markLastWasBlockQuote() {
        lastWasBlockQuote = true;
    }

    void markLastWasNotBlockQuote() {
        lastWasBlockQuote = false;
    }

    /**
     * 引用行とその描画情報を記録する。
     */
    void recordRow(int rowNum, int quoteDecorCol, RenderState.QuoteRowInfo rowInfo) {

        if (!open || firstRow < 0) {
            open = true;
            firstRow = rowNum;
            startCol = quoteDecorCol;
        }

        if (quoteDecorCol < startCol) {
            startCol = quoteDecorCol;
        }

        lastRow = rowNum;

        rows.put(Integer.valueOf(rowNum), rowInfo);

        lastWasBlockQuote = true;
    }

    RenderState.QuoteRowInfo getRowInfo(int rowNum) {
        return rows.get(Integer.valueOf(rowNum));
    }

    void replaceRowInfo(int rowNum, RenderState.QuoteRowInfo rowInfo) {

        rows.put(Integer.valueOf(rowNum), rowInfo);
    }

    void updateTableRowStyleRole(int rowNum, MarkdownTable.TableRowStyleRole role) {

        RenderState.QuoteRowInfo current = getRowInfo(rowNum);

        if (current == null || current.getKind() != RenderState.QuoteRowKind.TABLE) {
            return;
        }

        replaceRowInfo(rowNum, current.withTableRowStyleRole(role));
    }

    boolean isRowKind(int rowNum, RenderState.QuoteRowKind kind) {

        RenderState.QuoteRowInfo info = getRowInfo(rowNum);

        return info != null && info.getKind() == kind;
    }

    void clearRows() {
        rows.clear();
    }

    /**
     * 現在開いている引用範囲を破棄する。
     *
     * lastWasBlockQuoteは、引用直後の空行判定に使うため維持する。
     */
    void clearTracking() {
        open = false;
        firstRow = -1;
        lastRow = -1;
        startCol = 0;

        clearRows();
    }
}