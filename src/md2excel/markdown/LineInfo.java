package md2excel.markdown;

/**
 * Markdownの1行を解析した結果。
 *
 * 引用行の場合も再帰的なLineInfoを作らず、 quoteDepthと引用マーカー除去後のcontent情報を保持する。
 */
public final class LineInfo {

    // 入力Markdown上の行全体
    private final String raw;
    private final String trimmed;
    private final int indent;

    // 引用マーカーを除去した内容
    // 非引用行の場合はraw/trimmed/indentと同じ
    private final String contentRaw;
    private final String contentTrimmed;
    private final int contentIndent;

    private final int quoteDepth;
    private final LineKind kind;

    private final int headingLevel;
    private final String headingText;

    private final boolean endsWithHardBreak;
    private final String paragraphText;
    private final String listMarkerText;
    private final String listContentText;

    private LineInfo(String raw, String trimmed, int indent, String contentRaw, String contentTrimmed,
            int contentIndent, int quoteDepth, LineKind kind, int headingLevel, String headingText,
            boolean endsWithHardBreak, String paragraphText, String listMarkerText, String listContentText) {

        if (kind == null) {
            throw new IllegalArgumentException("kind must not be null");
        }

        if (quoteDepth < 0) {
            throw new IllegalArgumentException("quoteDepth must not be negative: " + quoteDepth);
        }

        this.raw = raw;
        this.trimmed = trimmed;
        this.indent = indent;

        this.contentRaw = contentRaw;
        this.contentTrimmed = contentTrimmed;
        this.contentIndent = contentIndent;

        this.quoteDepth = quoteDepth;
        this.kind = kind;

        this.headingLevel = headingLevel;
        this.headingText = headingText;

        this.endsWithHardBreak = endsWithHardBreak;
        this.paragraphText = paragraphText;
        this.listMarkerText = listMarkerText;
        this.listContentText = listContentText;
    }

    public static LineInfo codeFence(String raw, String trimmed, int indent) {

        return simple(raw, trimmed, indent, LineKind.CODE_FENCE);
    }

    public static LineInfo codeLine(String raw, String trimmed, int indent) {

        return simple(raw, trimmed, indent, LineKind.CODE_LINE);
    }

    public static LineInfo blank(String raw, String trimmed, int indent) {

        return simple(raw, trimmed, indent, LineKind.BLANK);
    }

    public static LineInfo horizontalRule(String raw, String trimmed, int indent) {

        return simple(raw, trimmed, indent, LineKind.HORIZONTAL_RULE);
    }

    public static LineInfo tableSeparator(String raw, String trimmed, int indent) {

        return simple(raw, trimmed, indent, LineKind.TABLE_SEPARATOR);
    }

    public static LineInfo tableRow(String raw, String trimmed, int indent) {

        return simple(raw, trimmed, indent, LineKind.TABLE_ROW);
    }

    public static LineInfo heading(String raw, String trimmed, int indent, int headingLevel, String headingText,
            boolean endsWithHardBreak) {

        return new LineInfo(raw, trimmed, indent, raw, trimmed, indent, 0, LineKind.HEADING, headingLevel, headingText,
                endsWithHardBreak, null, null, null);
    }

    public static LineInfo bulletItem(String raw, String trimmed, int indent, String markerText, String contentText,
            boolean endsWithHardBreak) {

        return listItem(raw, trimmed, indent, LineKind.BULLET_ITEM, markerText, contentText, endsWithHardBreak);
    }

    public static LineInfo numberItem(String raw, String trimmed, int indent, String markerText, String contentText,
            boolean endsWithHardBreak) {

        return listItem(raw, trimmed, indent, LineKind.NUMBER_ITEM, markerText, contentText, endsWithHardBreak);
    }

    public static LineInfo normal(String raw, String trimmed, int indent, String paragraphText,
            boolean endsWithHardBreak) {

        return new LineInfo(raw, trimmed, indent, raw, trimmed, indent, 0, LineKind.NORMAL, -1, null, endsWithHardBreak,
                paragraphText, null, null);
    }

    /**
     * 引用マーカー除去後に解析したLineInfoを、 明示的な引用情報付きLineInfoへ変換する。
     *
     * content自体は保持せず、値だけをコピーするため、 再帰的なLineInfo構造にはならない。
     */
    public static LineInfo quoted(String raw, String trimmed, int indent, int quoteDepth, LineInfo content) {

        if (quoteDepth <= 0) {
            throw new IllegalArgumentException("quoteDepth must be positive: " + quoteDepth);
        }

        if (content == null) {
            throw new IllegalArgumentException("content must not be null");
        }

        if (content.isQuoted()) {
            throw new IllegalArgumentException("content must already be unwrapped");
        }

        return new LineInfo(raw, trimmed, indent, content.getContentRaw(), content.getContentTrimmed(),
                content.getContentIndent(), quoteDepth, content.getKind(), content.getHeadingLevel(),
                content.getHeadingText(), content.endsWithHardBreak(), content.getParagraphText(),
                content.getListMarkerText(), content.getListContentText());
    }

    private static LineInfo simple(String raw, String trimmed, int indent, LineKind kind) {

        return new LineInfo(raw, trimmed, indent, raw, trimmed, indent, 0, kind, -1, null, false, null, null, null);
    }

    private static LineInfo listItem(String raw, String trimmed, int indent, LineKind kind, String markerText,
            String contentText, boolean endsWithHardBreak) {

        return new LineInfo(raw, trimmed, indent, raw, trimmed, indent, 0, kind, -1, null, endsWithHardBreak,
                contentText, markerText, contentText);
    }

    public String getRaw() {
        return raw;
    }

    public String getTrimmed() {
        return trimmed;
    }

    /**
     * 引用マーカーより前のインデント。 引用ブロックの配置計算に使用する。
     */
    public int getIndent() {
        return indent;
    }

    /**
     * 引用マーカーを除去した生テキスト。 非引用行の場合はgetRaw()と同じ。
     */
    public String getContentRaw() {
        return contentRaw;
    }

    /**
     * 引用マーカーを除去した内容のtrim結果。
     */
    public String getContentTrimmed() {
        return contentTrimmed;
    }

    /**
     * 引用マーカーを除去した内容のインデント。 引用内リストやコードブロックで使用する。
     */
    public int getContentIndent() {
        return contentIndent;
    }

    public LineKind getKind() {
        return kind;
    }

    public boolean isQuoted() {
        return quoteDepth > 0;
    }

    public int getQuoteDepth() {
        return quoteDepth;
    }

    public int getHeadingLevel() {
        return headingLevel;
    }

    public String getHeadingText() {
        return headingText;
    }

    public boolean endsWithHardBreak() {
        return endsWithHardBreak;
    }

    public String getParagraphText() {
        return paragraphText;
    }

    public String getListMarkerText() {
        return listMarkerText;
    }

    public String getListContentText() {
        return listContentText;
    }

    public boolean isTableLike() {
        return kind == LineKind.TABLE_SEPARATOR || kind == LineKind.TABLE_ROW;
    }
}