package md2excel.render;

/**
 * 引用行の装飾に必要な情報。
 */
final class QuoteRowInfo {

    private final QuoteRowKind kind;
    private final int depth;
    private final int contentCol;

    private final int tableStartCol;
    private final int tableEndCol;
    private final TableRowStyleRole tableRowStyleRole;

    QuoteRowInfo(QuoteRowKind kind, int depth, int contentCol) {
        this(kind, depth, contentCol, -1, -1, null);
    }

    QuoteRowInfo(QuoteRowKind kind, int depth, int contentCol, int tableStartCol, int tableEndCol, TableRowStyleRole tableRowStyleRole) {

        this.kind = kind;
        this.depth = Math.max(1, depth);
        this.contentCol = contentCol;
        this.tableStartCol = tableStartCol;
        this.tableEndCol = tableEndCol;
        this.tableRowStyleRole = tableRowStyleRole;
    }

    QuoteRowInfo withTableRowStyleRole(TableRowStyleRole role) {

        return new QuoteRowInfo(kind, depth, contentCol, tableStartCol, tableEndCol, role);
    }

    boolean isTableContentColumn(int col) {
        return tableStartCol >= 0 && tableEndCol >= tableStartCol && col >= tableStartCol && col <= tableEndCol;
    }

    QuoteRowKind getKind() {
        return kind;
    }

    int getDepth() {
        return depth;
    }

    int getContentCol() {
        return contentCol;
    }

    TableRowStyleRole getTableRowStyleRole() {
        return tableRowStyleRole;
    }
}