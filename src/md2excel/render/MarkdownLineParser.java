package md2excel.render;

import md2excel.markdown.CodeFence;
import md2excel.markdown.LineInfo;
import md2excel.markdown.MdTextUtil;
import md2excel.markdown.NumberedListMarker;

final class MarkdownLineParser {

    private MarkdownLineParser() {
    }

    static LineInfo parse(String rawLine, RenderState st, boolean tableParsingEnabled) {

        if (st.codeBlock().isOpen()) {
            if (st.codeBlock().isInBlockQuote()) {
                return parseQuotedCodeBlockLine(rawLine, st);
            }

            return parseCodeBlockContent(rawLine, st);
        }

        QuotePrefix quote = unwrapQuoteMarkers(rawLine, Integer.MAX_VALUE);

        if (quote.depth > 0) {
            LineInfo content = parseContent(quote.contentRaw, tableParsingEnabled);

            return LineInfo.quoted(rawLine, rawLine.trim(), MdTextUtil.countLeadingSpacesOrTabs(rawLine), quote.depth,
                    content);
        }

        return parseContent(rawLine, tableParsingEnabled);
    }

    private static LineInfo parseQuotedCodeBlockLine(String rawLine, RenderState st) {

        int openingQuoteDepth = st.codeBlock().getQuoteDepth();

        QuotePrefix quote = unwrapQuoteMarkers(rawLine, openingQuoteDepth);

        LineInfo content = parseCodeBlockContent(quote.contentRaw, st);

        if (quote.depth == 0) {
            return content;
        }

        return LineInfo.quoted(rawLine, rawLine.trim(), MdTextUtil.countLeadingSpacesOrTabs(rawLine), quote.depth,
                content);
    }

    private static LineInfo parseCodeBlockContent(String rawLine, RenderState st) {

        String trimmed = rawLine.trim();
        int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);

        if (CodeFence.isClosingLine(trimmed, st.codeBlock().getFenceMarker(), st.codeBlock().getFenceLength())) {

            return LineInfo.codeFence(rawLine, trimmed, indent);
        }

        return LineInfo.codeLine(rawLine, trimmed, indent);
    }

    private static LineInfo parseContent(String rawLine, boolean tableParsingEnabled) {

        String trimmed = rawLine.trim();
        int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);
        boolean endsWithHardBreak = hasLineEndHardBreak(rawLine);

        // code fence
        if (CodeFence.parseOpening(trimmed) != null) {
            return LineInfo.codeFence(rawLine, trimmed, indent);
        }

        // blank
        if (trimmed.isEmpty()) {
            return LineInfo.blank(rawLine, trimmed, indent);
        }

        // horizontal rule
        if (MdTextUtil.isHorizontalRuleLine(trimmed)) {
            return LineInfo.horizontalRule(rawLine, trimmed, indent);
        }

        // table
        if (tableParsingEnabled && MarkdownTable.isTableLine(rawLine)) {

            boolean separator = MarkdownTable.isTableSeparatorLine(trimmed);

            return separator ? LineInfo.tableSeparator(rawLine, trimmed, indent)
                    : LineInfo.tableRow(rawLine, trimmed, indent);
        }

        // heading
        if (trimmed.startsWith("#")) {
            int level = MdTextUtil.countHeadingLevel(trimmed);

            String text = trimmed.substring(level).trim();
            text = MdTextUtil.stripHeadingClosingHashes(text);
            text = stripLineEndHardBreakMarker(text, rawLine);

            return LineInfo.heading(rawLine, trimmed, indent, level, text, endsWithHardBreak);
        }

        // bullet list
        if (trimmed.length() >= 2) {
            char marker = trimmed.charAt(0);

            if ((marker == '*' || marker == '-' || marker == '+') && Character.isWhitespace(trimmed.charAt(1))) {

                String content = trimmed.substring(2).trim();
                content = stripLineEndHardBreakMarker(content, rawLine);

                return LineInfo.bulletItem(rawLine, trimmed, indent, "・ ", content, endsWithHardBreak);
            }
        }

        // numbered list
        NumberedListMarker numberedMarker = NumberedListMarker.parse(trimmed);

        if (numberedMarker != null) {
            String content = trimmed.substring(numberedMarker.getContentStartIndex()).trim();

            content = stripLineEndHardBreakMarker(content, rawLine);

            return LineInfo.numberItem(rawLine, trimmed, indent, numberedMarker.getMarkerText(), content,
                    endsWithHardBreak);
        }

        // normal paragraph
        String paragraphText = stripLineEndHardBreakMarker(trimmed, rawLine);

        return LineInfo.normal(rawLine, trimmed, indent, paragraphText, endsWithHardBreak);
    }

    static String stripOneQuoteMarker(String rawLine) {
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

    private static boolean hasLineEndHardBreak(String rawLine) {
        return MdTextUtil.hasHardLineBreakByBackslash(rawLine) || MdTextUtil.hasHardLineBreakBySpaces(rawLine);
    }

    private static String stripLineEndHardBreakMarker(String text, String rawLine) {

        if (text == null) {
            return null;
        }

        if (MdTextUtil.hasHardLineBreakByBackslash(rawLine)) {
            return MdTextUtil.removeTrailingBackslash(text);
        }

        return text;
    }

    private static final class QuotePrefix {

        final int depth;
        final String contentRaw;

        QuotePrefix(int depth, String contentRaw) {

            this.depth = depth;
            this.contentRaw = contentRaw;
        }
    }

    private static QuotePrefix unwrapQuoteMarkers(String rawLine, int maxDepth) {

        int depth = 0;
        String content = rawLine;

        while (depth < maxDepth && content.trim().startsWith(">")) {

            content = stripOneQuoteMarker(content);
            depth++;
        }

        return new QuotePrefix(depth, content);
    }
}