package md2excel.render;

/**
 * 現在行と次の行から、現在行を表として解析するか判定する。
 */
final class TableParsingProbe {

    private TableParsingProbe() {
    }

    static boolean shouldEnableTableParsing(String rawLine, String nextRawLine, RenderState state) {

        QuotePrefix current = QuotePrefix.unwrap(rawLine, Integer.MAX_VALUE);

        boolean continuingTable = state.isLastLineTable() && state.table().getQuoteDepth() == current.depth;

        if (continuingTable) {
            return true;
        }

        if (nextRawLine == null) {
            return false;
        }

        QuotePrefix next = QuotePrefix.unwrap(nextRawLine, Integer.MAX_VALUE);

        return current.depth == next.depth && MarkdownTable.isTableStart(current.contentRaw, next.contentRaw);
    }
}