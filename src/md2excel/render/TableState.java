package md2excel.render;

final class TableState {

    private int startCol;
    private int headerRow = -1;
    private int bodyStartRow = -1;
    private int lastBodyRow = -1;
    private int endCol = -1;
    private int quoteDepth = -1;

    boolean isOpen() {
        return headerRow >= 0;
    }

    int getStartCol() {
        return startCol;
    }

    int getHeaderRow() {
        return headerRow;
    }

    int getBodyStartRow() {
        return bodyStartRow;
    }

    int getLastBodyRow() {
        return lastBodyRow;
    }

    int getEndCol() {
        return endCol;
    }

    int getQuoteDepth() {
        return quoteDepth;
    }

    void begin(int headerRow, int startCol, int endCol, int quoteDepth) {

        this.headerRow = headerRow;
        this.startCol = startCol;
        this.endCol = endCol;
        this.quoteDepth = quoteDepth;

        this.bodyStartRow = -1;
        this.lastBodyRow = -1;
    }

    void recordBodyRows(int firstRow, int lastRow, int lastCol) {

        if (bodyStartRow < 0) {
            bodyStartRow = firstRow;
        }

        lastBodyRow = lastRow;
        endCol = Math.max(endCol, lastCol);
    }

    void reset() {
        startCol = 0;
        headerRow = -1;
        bodyStartRow = -1;
        lastBodyRow = -1;
        endCol = -1;
        quoteDepth = -1;
    }
}