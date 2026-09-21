package md2excel.markdown;

public final class MdTextUtil {
    private static final String TAB_SPACES = "    ";

    private MdTextUtil() {
    }

    public static String expandTabs(String s) {
        if (s == null || s.isEmpty()) {
            return s;
        }
        return s.replace("\t", TAB_SPACES);
    }

    public static int countLeadingSpacesOrTabs(String s) {
        int count = 0;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch == ' ')
                count++;
            else if (ch == '\t')
                count += 4;
            else
                break;
        }
        return count;
    }

    public static int countHeadingLevel(String trimmedLine) {
        int count = 0;
        for (int i = 0; i < trimmedLine.length(); i++) {
            if (trimmedLine.charAt(i) == '#')
                count++;
            else
                break;
        }
        return count;
    }

    public static boolean isAsciiLike(char ch) {
        return ch >= 0x20 && ch <= 0x7E;
    }

    // 行末の半角スペース2個以上でハード改行
    public static boolean hasHardLineBreakBySpaces(String rawLine) {
        if (rawLine == null || rawLine.isEmpty()) {
            return false;
        }
        int count = 0;
        for (int i = rawLine.length() - 1; i >= 0; i--) {
            char ch = rawLine.charAt(i);
            if (ch == ' ') {
                count++;
                if (count >= 2) {
                    return true;
                }
                continue;
            }
            break;
        }
        return false;
    }

    // 行末のバックスラッシュでハード改行（末尾の空白/タブは無視）
    public static boolean hasHardLineBreakByBackslash(String rawLine) {
        if (rawLine == null || rawLine.isEmpty()) {
            return false;
        }
        int end = rawLine.length();
        while (end > 0) {
            char ch = rawLine.charAt(end - 1);
            if (ch == ' ' || ch == '\t') {
                end--;
                continue;
            }
            return ch == '\\';
        }
        return false;
    }

    public static String removeTrailingBackslash(String text) {
        if (text == null || text.isEmpty()) {
            return text;
        }
        if (text.charAt(text.length() - 1) == '\\') {
            return text.substring(0, text.length() - 1).trim();
        }
        return text;
    }

    // "## title ##" のような閉じ # を取り除く（末尾は空白のみでもOK）
    public static String stripHeadingClosingHashes(String text) {
        if (text == null || text.isEmpty()) {
            return text;
        }

        int end = text.length();
        while (end > 0 && Character.isWhitespace(text.charAt(end - 1))) {
            end--;
        }
        if (end == 0) {
            return "";
        }

        int hashEnd = end;
        while (hashEnd > 0 && text.charAt(hashEnd - 1) == '#') {
            hashEnd--;
        }
        if (hashEnd == end) {
            return text;
        }

        if (hashEnd == 0 || !Character.isWhitespace(text.charAt(hashEnd - 1))) {
            return text;
        }

        int trimEnd = hashEnd - 1;
        while (trimEnd > 0 && Character.isWhitespace(text.charAt(trimEnd - 1))) {
            trimEnd--;
        }
        return text.substring(0, trimEnd);
    }

    public static String replaceBrOutsideInlineCode(String s, String replacement) {

        return MdInlineCodeUtil.replaceBrOutsideCodeSpans(s, replacement);
    }

    // "---", "***", "___", "- - -" のような水平線を判定（空白/タブのみ許可）
    public static boolean isHorizontalRuleLine(String trimmed) {
        if (trimmed == null || trimmed.isEmpty()) {
            return false;
        }

        char marker = '\0';
        int count = 0;
        for (int i = 0; i < trimmed.length(); i++) {
            char ch = trimmed.charAt(i);
            if (ch == ' ' || ch == '\t') {
                continue;
            }
            if (ch != '-' && ch != '_' && ch != '*') {
                return false;
            }
            if (marker == '\0') {
                marker = ch;
            } else if (marker != ch) {
                return false;
            }
            count++;
        }
        return count >= 3;
    }

    public static String removeLeadingIndentColumns(String s, int columnsToRemove) {
        if (s == null || s.isEmpty() || columnsToRemove <= 0) {
            return s;
        }

        int index = 0;
        int removedColumns = 0;

        while (index < s.length() && removedColumns < columnsToRemove) {
            char ch = s.charAt(index);

            if (ch != ' ' && ch != '\t') {
                break;
            }

            int width = (ch == '\t') ? 4 : 1;
            int remaining = columnsToRemove - removedColumns;

            // タブの一部だけを除去する場合
            if (remaining < width) {
                int remainingSpaces = width - remaining;
                StringBuilder result = new StringBuilder();

                for (int i = 0; i < remainingSpaces; i++) {
                    result.append(' ');
                }

                result.append(s, index + 1, s.length());
                return result.toString();
            }

            removedColumns += width;
            index++;
        }

        return s.substring(index);
    }
}