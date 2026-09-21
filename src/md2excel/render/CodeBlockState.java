package md2excel.render;

final class CodeBlockState {

    private boolean open;

    private char fenceMarker;
    private int fenceLength;
    private int openingIndent;

    private int firstRow = -1;
    private int lastRow = -1;
    private int startCol;

    private boolean inBlockQuote;
    private int quoteStartCol = -1;
    private int quoteDepth;

    boolean isOpen() {
        return open;
    }

    char getFenceMarker() {
        return fenceMarker;
    }

    int getFenceLength() {
        return fenceLength;
    }

    int getOpeningIndent() {
        return openingIndent;
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

    boolean isInBlockQuote() {
        return inBlockQuote;
    }

    int getQuoteStartCol() {
        return quoteStartCol;
    }

    int getQuoteDepth() {
        return quoteDepth;
    }

    boolean hasRenderedLines() {
        return firstRow >= 0 && lastRow >= 0;
    }

    void open(char fenceMarker, int fenceLength, int openingIndent, boolean inBlockQuote, int quoteStartCol,
            int quoteDepth) {

        this.open = true;
        this.fenceMarker = fenceMarker;
        this.fenceLength = fenceLength;
        this.openingIndent = openingIndent;

        this.firstRow = -1;
        this.lastRow = -1;
        this.startCol = 0;

        this.inBlockQuote = inBlockQuote;
        this.quoteStartCol = inBlockQuote ? quoteStartCol : -1;

        this.quoteDepth = inBlockQuote ? Math.max(1, quoteDepth) : 0;
    }

    void recordLine(int rowNum, int startCol) {
        if (firstRow < 0) {
            firstRow = rowNum;
            this.startCol = startCol;
        }

        lastRow = rowNum;
    }

    void reset() {
        open = false;

        fenceMarker = '\0';
        fenceLength = 0;
        openingIndent = 0;

        firstRow = -1;
        lastRow = -1;
        startCol = 0;

        inBlockQuote = false;
        quoteStartCol = -1;
        quoteDepth = 0;
    }
}