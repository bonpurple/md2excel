package md2excel.render;

/**
 * Setext見出しの下線を判定する。
 *
 * 描画状態やApache POIには依存しない。
 */
final class SetextHeadingDetector {

    private SetextHeadingDetector() {
    }

    /**
     * @return 見出しレベル。該当しなければ0
     */
    static int detect(String raw, boolean hasParagraphContent, boolean normalParagraph, int baseIndent) {

        if (!hasParagraphContent || !normalParagraph || baseIndent > 3) {
            return 0;
        }

        return parseUnderlineLevel(raw);
    }

    private static int parseUnderlineLevel(String raw) {
        if (raw == null || raw.isEmpty()) {
            return 0;
        }

        int n = raw.length();
        int i = 0;
        int leadingSpaces = 0;

        while (i < n && raw.charAt(i) == ' ') {
            leadingSpaces++;
            if (leadingSpaces > 3) {
                return 0;
            }
            i++;
        }

        // marker 前の tab は4-column境界になるので、
        // この簡易判定では Setext underline としない。
        if (i < n && raw.charAt(i) == '\t') {
            return 0;
        }

        if (i >= n) {
            return 0;
        }

        char marker = raw.charAt(i);
        if (marker != '=' && marker != '-') {
            return 0;
        }

        while (i < n && raw.charAt(i) == marker) {
            i++;
        }

        // marker の後ろは space / tab のみ許可。
        while (i < n) {
            char ch = raw.charAt(i);
            if (ch != ' ' && ch != '\t') {
                return 0;
            }
            i++;
        }

        return marker == '=' ? 1 : 2;
    }
}