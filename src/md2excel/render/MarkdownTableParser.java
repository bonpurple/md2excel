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

        List<Integer> pipePositions = new ArrayList<Integer>();
        scanUnescapedPipes(inner, pipePositions);

        List<String> cells = new ArrayList<String>(pipePositions.size() + 1);
        int segStart = 0;

        for (int pipePosition : pipePositions) {
            cells.add(inner.substring(segStart, pipePosition));
            segStart = pipePosition + 1;
        }

        // 空の最終セルも、従来どおり追加する。
        cells.add(inner.substring(segStart));

        return cells;
    }

    private static int countUnescapedPipes(String s) {
        if (s == null || s.isEmpty()) {
            return 0;
        }

        return scanUnescapedPipes(s, null);
    }

    /**
     * エスケープされていないパイプを左から走査する。
     *
     * 直前に連続するバックスラッシュが偶数個なら区切りとする。 positionsがnullの場合は個数だけ数え、セル分割時には位置も記録する。
     */
    private static int scanUnescapedPipes(String text, List<Integer> positions) {
        int count = 0;
        int consecutiveBackslashes = 0;

        for (int index = 0; index < text.length(); index++) {
            char ch = text.charAt(index);

            if (ch == '\\') {
                consecutiveBackslashes++;
                continue;
            }

            if (ch == '|' && (consecutiveBackslashes % 2) == 0) {
                count++;

                if (positions != null) {
                    positions.add(Integer.valueOf(index));
                }
            }

            consecutiveBackslashes = 0;
        }

        return count;
    }

    /**
     * テーブルセル内の隣接した "\|" を "|" に置換する。 インラインコード内も区別せず、従来どおり同じ処理を行う。
     */
    static String unescapePipesInTableCell(String s) {
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
