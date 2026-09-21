package md2excel.markdown;

/**
 * fenced code blockの開始フェンス情報。
 */
public final class CodeFence {

    private final char marker;
    private final int length;

    private CodeFence(char marker, int length) {
        this.marker = marker;
        this.length = length;
    }

    /**
     * 開始フェンスを解析する。
     *
     * 開始フェンスとして扱う条件: - 先頭がバッククォートまたはチルダ - 同じ記号が3文字以上連続
     *
     * 既存仕様を維持し、フェンス後方の文字列は検証しない。
     */
    public static CodeFence parseOpening(String trimmedLine) {
        if (trimmedLine == null || trimmedLine.isEmpty()) {

            return null;
        }

        char marker = trimmedLine.charAt(0);

        if (marker != '`' && marker != '~') {
            return null;
        }

        int length = countMarkerRun(trimmedLine, marker);

        if (length < 3) {
            return null;
        }

        return new CodeFence(marker, length);
    }

    /**
     * 閉じフェンスか判定する。
     *
     * 条件: - 開始フェンスと同じ記号 - 開始フェンス以上の長さ - フェンス後方は空白のみ
     */
    public static boolean isClosingLine(String trimmedLine, char openingMarker, int openingLength) {

        if (trimmedLine == null || trimmedLine.isEmpty()) {

            return false;
        }

        if (openingMarker != '`' && openingMarker != '~') {

            return false;
        }

        if (openingLength < 3) {
            return false;
        }

        if (trimmedLine.charAt(0) != openingMarker) {
            return false;
        }

        int length = countMarkerRun(trimmedLine, openingMarker);

        if (length < openingLength) {
            return false;
        }

        for (int index = length; index < trimmedLine.length(); index++) {

            if (!Character.isWhitespace(trimmedLine.charAt(index))) {

                return false;
            }
        }

        return true;
    }

    public char getMarker() {
        return marker;
    }

    public int getLength() {
        return length;
    }

    private static int countMarkerRun(String text, char marker) {

        int length = 0;

        while (length < text.length() && text.charAt(length) == marker) {

            length++;
        }

        return length;
    }
}