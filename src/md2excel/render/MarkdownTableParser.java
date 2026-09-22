package md2excel.render;

import java.util.ArrayList;
import java.util.List;

final class MarkdownTableParser {

    private MarkdownTableParser() {
    }

    static boolean isTableLine(String line) {
        String trimmed = line.trim();
        return countUnescapedPipes(trimmed) >= 1;
    }

    static boolean isTableSeparatorLine(String line) {
        if (line == null) {
            return false;
        }

        String trimmed = line.trim();
        if (countUnescapedPipes(trimmed) < 1) {
            return false;
        }

        List<String> cells = splitTableCells(trimmed);
        if (cells.isEmpty()) {
            return false;
        }

        for (String cell : cells) {
            if (!isTableSeparatorCell(cell)) {
                return false;
            }
        }

        return true;
    }

    static boolean isTableStart(String headerLine, String separatorLine) {
        if (!isTableLine(headerLine) || isTableSeparatorLine(headerLine) || !isTableSeparatorLine(separatorLine)) {
            return false;
        }

        return splitTableCells(headerLine).size() == splitTableCells(separatorLine).size();
    }

    private static boolean isTableSeparatorCell(String cell) {
        StringBuilder compact = new StringBuilder();

        for (int i = 0; i < cell.length(); i++) {
            char ch = cell.charAt(i);
            if (!Character.isWhitespace(ch)) {
                compact.append(ch);
            }
        }

        int start = 0;
        int end = compact.length();

        if (start < end && compact.charAt(start) == ':') {
            start++;
        }
        if (start < end && compact.charAt(end - 1) == ':') {
            end--;
        }

        int hyphenCount = 0;
        for (int i = start; i < end; i++) {
            if (compact.charAt(i) != '-') {
                return false;
            }
            hyphenCount++;
        }

        // 既存の "| -- | -- |" も許容する
        return hyphenCount >= 1;
    }

    static List<String> splitTableCells(String line) {
        String trimmed = line.trim();
        String inner = trimmed;

        if (inner.startsWith("|")) {
            inner = inner.substring(1);
        }
        if (inner.endsWith("|")) {
            inner = inner.substring(0, inner.length() - 1);
        }

        List<String> cells = new ArrayList<String>();

        int segStart = 0;
        int n = inner.length();

        for (int i = 0; i <= n; i++) {
            if (i == n || (inner.charAt(i) == '|' && !isEscapedPipe(inner, i))) {
                cells.add(inner.substring(segStart, i));
                segStart = i + 1;
            }
        }

        return cells;
    }

    /**
     * pos の '|' が "\|" のようにエスケープされているか判定する。 直前に連続する '\' の個数が奇数ならエスケープ扱い。
     */
    private static boolean isEscapedPipe(String s, int pos) {
        if (pos <= 0 || pos >= s.length() || s.charAt(pos) != '|')
            return false;
        int bs = 0;
        for (int i = pos - 1; i >= 0 && s.charAt(i) == '\\'; i--) {
            bs++;
        }
        return (bs % 2) == 1;
    }

    private static int countUnescapedPipes(String s) {
        if (s == null || s.isEmpty())
            return 0;

        int count = 0;
        for (int i = 0; i < s.length(); i++) {
            if (s.charAt(i) == '|' && !isEscapedPipe(s, i)) {
                count++;
            }
        }
        return count;
    }

    /**
     * テーブルセル内の "\|" を "|" に戻す。
     */
    static String unescapePipeOutsideInlineCode(String s) {
        if (s == null || s.isEmpty())
            return s;
        StringBuilder out = new StringBuilder(s.length());
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch == '\\' && i + 1 < s.length() && s.charAt(i + 1) == '|') {
                out.append('|');
                i++; // '|' を消費
                continue;
            }
            out.append(ch);
        }
        return out.toString();
    }

}
