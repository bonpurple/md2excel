package md2excel.markdown;

public final class MdInlineCodeUtil {

    private MdInlineCodeUtil() {
    }

    public static int countBackticks(String text, int position) {
        int count = 0;

        while (position + count < text.length() && text.charAt(position + count) == '`') {
            count++;
        }

        return count;
    }

    public static int findClosingBackticks(String text, int start, int tickLength) {

        for (int index = start; index < text.length();) {
            if (text.charAt(index) != '`') {
                index++;
                continue;
            }

            int runLength = countBackticks(text, index);

            if (runLength == tickLength) {
                return index;
            }

            index += runLength;
        }

        return -1;
    }

    public static String replaceBrOutsideCodeSpans(String text, String replacement) {

        if (text == null || text.isEmpty()) {
            return text;
        }

        StringBuilder out = new StringBuilder(text.length());

        for (int index = 0; index < text.length();) {
            char ch = text.charAt(index);

            // エスケープされた記号は構文として扱わない
            if (ch == '\\' && index + 1 < text.length() && isAsciiPunctuation(text.charAt(index + 1))) {

                out.append(ch);
                out.append(text.charAt(index + 1));
                index += 2;
                continue;
            }

            // 複数長バッククォートを含むcode span
            if (ch == '`') {
                int tickLength = countBackticks(text, index);

                int closingPosition = findClosingBackticks(text, index + tickLength, tickLength);

                if (closingPosition >= 0) {
                    int end = closingPosition + tickLength;

                    out.append(text, index, end);
                    index = end;
                    continue;
                }

                // 閉じrunがなければ通常文字として残す
                out.append(text, index, index + tickLength);
                index += tickLength;
                continue;
            }

            int brLength = matchBrTagLength(text, index);

            if (brLength > 0) {
                out.append(replacement);
                index += brLength;
                continue;
            }

            out.append(ch);
            index++;
        }

        return out.toString();
    }

    public static int matchBrTagLength(String text, int position) {
        int length = text.length();

        if (position < 0 || position + 3 >= length) {
            return 0;
        }

        if (text.charAt(position) != '<') {
            return 0;
        }

        char b = text.charAt(position + 1);
        char r = text.charAt(position + 2);

        if (Character.toLowerCase(b) != 'b' || Character.toLowerCase(r) != 'r') {
            return 0;
        }

        int index = position + 3;

        while (index < length && Character.isWhitespace(text.charAt(index))) {
            index++;
        }

        if (index < length && text.charAt(index) == '/') {
            index++;

            while (index < length && Character.isWhitespace(text.charAt(index))) {
                index++;
            }
        }

        if (index < length && text.charAt(index) == '>') {
            return index - position + 1;
        }

        return 0;
    }

    private static boolean isAsciiPunctuation(char ch) {
        if (ch > 0x7F) {
            return false;
        }

        return (ch >= '!' && ch <= '/') || (ch >= ':' && ch <= '@') || (ch >= '[' && ch <= '`')
                || (ch >= '{' && ch <= '~');
    }
}