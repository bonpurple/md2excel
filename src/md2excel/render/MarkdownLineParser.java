package md2excel.render;

import md2excel.markdown.MdTextUtil;
import md2excel.render.MarkdownRenderer.LineInfo;
import md2excel.render.MarkdownRenderer.LineKind;

final class MarkdownLineParser {

    private MarkdownLineParser() {
    }

    static LineInfo parse(String rawLine, RenderState st, boolean tableParsingEnabled) {

        String trimmed = rawLine.trim();
        int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);

        // コードブロック中は最優先
        if (st.codeBlock().isOpen()) {
            if (st.codeBlock().isInBlockQuote() && trimmed.startsWith(">")) {

                String innerRaw = stripOneQuoteMarker(rawLine);
                LineInfo inner = parseCodeBlockContent(innerRaw, st);

                return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null, inner.endsWithHardBreak,
                        null, null, null, inner);
            }

            return parseCodeBlockContent(rawLine, st);
        }

        // 引用は外側のコンテキストとして扱う
        if (trimmed.startsWith(">")) {
            String innerRaw = stripOneQuoteMarker(rawLine);
            LineInfo inner = parseContent(innerRaw, tableParsingEnabled);

            return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null, inner.endsWithHardBreak, null,
                    null, null, inner);
        }

        return parseContent(rawLine, tableParsingEnabled);
    }

    private static LineInfo parseCodeBlockContent(String rawLine, RenderState st) {

        String trimmed = rawLine.trim();
        int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);

        if (MdTextUtil.isClosingCodeFenceLine(trimmed, st.codeBlock().getFenceMarker(),
                st.codeBlock().getFenceLength())) {

            return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_FENCE, -1, null, false, null, null, null, null);
        }

        return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_LINE, -1, null, false, null, null, null, null);
    }

    private static LineInfo parseContent(String rawLine, boolean tableParsingEnabled) {

        String trimmed = rawLine.trim();
        int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);
        boolean endsWithHardBreak = hasLineEndHardBreak(rawLine);

        // nested quote
        if (trimmed.startsWith(">")) {
            String innerRaw = stripOneQuoteMarker(rawLine);
            LineInfo inner = parseContent(innerRaw, tableParsingEnabled);

            return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null, inner.endsWithHardBreak, null,
                    null, null, inner);
        }

        // code fence
        if (MdTextUtil.isOpeningCodeFenceLine(trimmed)) {
            return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_FENCE, -1, null, false, null, null, null, null);
        }

        // blank
        if (trimmed.isEmpty()) {
            return new LineInfo(rawLine, trimmed, indent, LineKind.BLANK, -1, null, false, null, null, null, null);
        }

        // horizontal rule
        if (MdTextUtil.isHorizontalRuleLine(trimmed)) {
            return new LineInfo(rawLine, trimmed, indent, LineKind.HORIZONTAL_RULE, -1, null, false, null, null, null,
                    null);
        }

        // table
        if (tableParsingEnabled && MarkdownTable.isTableLine(rawLine)) {

            boolean separator = MarkdownTable.isTableSeparatorLine(trimmed);

            return new LineInfo(rawLine, trimmed, indent, separator ? LineKind.TABLE_SEPARATOR : LineKind.TABLE_ROW, -1,
                    null, false, null, null, null, null);
        }

        // heading
        if (trimmed.startsWith("#")) {
            int level = MdTextUtil.countHeadingLevel(trimmed);

            String text = trimmed.substring(level).trim();
            text = MdTextUtil.stripHeadingClosingHashes(text);
            text = stripLineEndHardBreakMarker(text, rawLine);

            return new LineInfo(rawLine, trimmed, indent, LineKind.HEADING, level, text, endsWithHardBreak, null, null,
                    null, null);
        }

        // bullet list
        if (trimmed.length() >= 2) {
            char marker = trimmed.charAt(0);

            if ((marker == '*' || marker == '-' || marker == '+') && Character.isWhitespace(trimmed.charAt(1))) {

                String content = trimmed.substring(2).trim();
                content = stripLineEndHardBreakMarker(content, rawLine);

                return new LineInfo(rawLine, trimmed, indent, LineKind.BULLET_ITEM, -1, null, endsWithHardBreak,
                        content, "・ ", content, null);
            }
        }

        // numbered list
        if (MdTextUtil.isNumberedListLine(trimmed)) {
            int markerEnd = findNumberedListMarkerEnd(trimmed);

            String markerText = trimmed.substring(0, markerEnd).trim() + " ";

            String content = trimmed.substring(markerEnd).trim();
            content = stripLineEndHardBreakMarker(content, rawLine);

            return new LineInfo(rawLine, trimmed, indent, LineKind.NUMBER_ITEM, -1, null, endsWithHardBreak, content,
                    markerText, content, null);
        }

        // normal paragraph
        String paragraphText = stripLineEndHardBreakMarker(trimmed, rawLine);

        return new LineInfo(rawLine, trimmed, indent, LineKind.NORMAL, -1, null, endsWithHardBreak, paragraphText, null,
                null, null);
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

    private static int findNumberedListMarkerEnd(String trimmed) {

        if (trimmed == null || trimmed.isEmpty()) {
            return -1;
        }

        int length = trimmed.length();
        int index = 0;

        while (index < length) {
            char ch = trimmed.charAt(index);

            if (ch < '0' || ch > '9') {
                break;
            }

            index++;
        }

        if (index == 0 || index >= length) {
            return -1;
        }

        char marker = trimmed.charAt(index);
        if (marker != '.' && marker != ')') {
            return -1;
        }

        index++;

        if (index >= length || !Character.isWhitespace(trimmed.charAt(index))) {
            return -1;
        }

        while (index < length && Character.isWhitespace(trimmed.charAt(index))) {
            index++;
        }

        return index;
    }
}