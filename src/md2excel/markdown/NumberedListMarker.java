package md2excel.markdown;

/**
 * 番号付きリストマーカーの解析結果。
 *
 * 例: "12.\ttext"
 *
 * markerText = "12. " contentStartIndex = 4
 */
public final class NumberedListMarker {

    private final String markerText;
    private final int contentStartIndex;

    private NumberedListMarker(String markerText, int contentStartIndex) {

        this.markerText = markerText;
        this.contentStartIndex = contentStartIndex;
    }

    /**
     * 行頭の番号付きリストマーカーを解析する。
     *
     * 対応形式: - "1. text" - "12.\ttext" - "1) text"
     *
     * @param text
     *            行頭空白を除去済みの文字列
     * @return 解析結果。番号付きリストでなければnull
     */
    public static NumberedListMarker parse(String text) {
        if (text == null || text.isEmpty()) {
            return null;
        }

        int length = text.length();
        int index = 0;

        char first = text.charAt(0);

        if (first < '0' || first > '9') {
            return null;
        }

        while (index < length) {
            char ch = text.charAt(index);

            if (ch < '0' || ch > '9') {
                break;
            }

            index++;
        }

        if (index >= length) {
            return null;
        }

        char punctuation = text.charAt(index);

        if (punctuation != '.' && punctuation != ')') {
            return null;
        }

        // 句読点の直後
        int markerEnd = index + 1;

        if (markerEnd >= length || !Character.isWhitespace(text.charAt(markerEnd))) {

            return null;
        }

        index = markerEnd;

        while (index < length && Character.isWhitespace(text.charAt(index))) {

            index++;
        }

        String markerText = text.substring(0, markerEnd) + " ";

        return new NumberedListMarker(markerText, index);
    }

    /**
     * Excelに表示する、空白を1個に正規化したマーカー。
     */
    public String getMarkerText() {
        return markerText;
    }

    /**
     * 元文字列内の本文開始位置。
     */
    public int getContentStartIndex() {
        return contentStartIndex;
    }
}