package md2excel.render;

/**
 * 行頭の引用マーカーを除去した結果。
 */
final class QuotePrefix {

    final int depth;
    final String contentRaw;

    private QuotePrefix(int depth, String contentRaw) {
        this.depth = depth;
        this.contentRaw = contentRaw;
    }

    static QuotePrefix unwrap(String rawLine, int maxDepth) {
        int depth = 0;
        String content = rawLine;

        while (depth < maxDepth && content.trim().startsWith(">")) {
            content = stripOneQuoteMarker(content);
            depth++;
        }

        return new QuotePrefix(depth, content);
    }

    private static String stripOneQuoteMarker(String rawLine) {
        int index = 0;

        while (index < rawLine.length()) {
            char ch = rawLine.charAt(index);

            if (ch == ' ' || ch == '\t') {
                index++;
                continue;
            }

            break;
        }

        if (index < rawLine.length() && rawLine.charAt(index) == '>') {
            index++;
        }

        if (index < rawLine.length() && rawLine.charAt(index) == ' ') {
            index++;
        }

        return index < rawLine.length() ? rawLine.substring(index) : "";
    }
}