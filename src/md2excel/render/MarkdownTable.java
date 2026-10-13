package md2excel.render;

/**
 * 表の判定に関する公開入口。
 */
final class MarkdownTable {

    private MarkdownTable() {
    }

    static boolean isTableLine(String line) {
        return MarkdownTableParser.isTableLine(line);
    }

    static boolean isTableSeparatorLine(String line) {
        return MarkdownTableParser.isTableSeparatorLine(line);
    }

    static boolean isTableStart(String headerLine, String separatorLine) {
        return MarkdownTableParser.isTableStart(headerLine, separatorLine);
    }
}