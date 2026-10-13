package md2excel.markdown;

/**
 * Markdownの1行を解析した結果。
 *
 * 引用行の場合も再帰的なLineInfoを作らず、quoteDepthと 引用マーカー除去後のcontent情報を保持する。
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

    /**
     * kindに応じて、HeadingDetail、ParagraphDetail、ListDetail、 または詳細なし（null）のいずれかを保持する。
     */
    private final Detail detail;

    private LineInfo(String raw, String trimmed, int indent, String contentRaw, String contentTrimmed, int contentIndent, int quoteDepth, LineKind kind,
        Detail detail) {

        if (kind == null) {
            throw new IllegalArgumentException("kind must not be null");
        }

        if (quoteDepth < 0) {
            throw new IllegalArgumentException("quoteDepth must not be negative: " + quoteDepth);
        }

        validateDetail(kind, detail);

        this.raw = raw;
        this.trimmed = trimmed;
        this.indent = indent;

        this.contentRaw = contentRaw;
        this.contentTrimmed = contentTrimmed;
        this.contentIndent = contentIndent;

        this.quoteDepth = quoteDepth;
        this.kind = kind;
        this.detail = detail;
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

    public static LineInfo heading(String raw, String trimmed, int indent, int headingLevel, String headingText, boolean endsWithHardBreak) {

        return unquoted(raw, trimmed, indent, LineKind.HEADING, new HeadingDetail(headingLevel, headingText, endsWithHardBreak));
    }

    public static LineInfo bulletItem(String raw, String trimmed, int indent, String markerText, String contentText, boolean endsWithHardBreak) {

        return listItem(raw, trimmed, indent, LineKind.BULLET_ITEM, markerText, contentText, endsWithHardBreak);
    }

    public static LineInfo numberItem(String raw, String trimmed, int indent, String markerText, String contentText, boolean endsWithHardBreak) {

        return listItem(raw, trimmed, indent, LineKind.NUMBER_ITEM, markerText, contentText, endsWithHardBreak);
    }

    public static LineInfo normal(String raw, String trimmed, int indent, String paragraphText, boolean endsWithHardBreak) {

        return unquoted(raw, trimmed, indent, LineKind.NORMAL, new ParagraphDetail(paragraphText, endsWithHardBreak));
    }

    /**
     * 引用マーカー除去後に解析したLineInfoを、 明示的な引用情報付きLineInfoへ変換する。
     *
     * 不変の詳細データを共有するが、content自体は保持しない。
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

        return new LineInfo(raw, trimmed, indent, content.contentRaw, content.contentTrimmed, content.contentIndent, quoteDepth, content.kind, content.detail);
    }

    private static LineInfo simple(String raw, String trimmed, int indent, LineKind kind) {

        return unquoted(raw, trimmed, indent, kind, null);
    }

    private static LineInfo listItem(String raw, String trimmed, int indent, LineKind kind, String markerText, String contentText, boolean endsWithHardBreak) {

        return unquoted(raw, trimmed, indent, kind, new ListDetail(markerText, contentText, endsWithHardBreak));
    }

    private static LineInfo unquoted(String raw, String trimmed, int indent, LineKind kind, Detail detail) {

        return new LineInfo(raw, trimmed, indent, raw, trimmed, indent, 0, kind, detail);
    }

    private static void validateDetail(LineKind kind, Detail detail) {
        boolean valid;

        switch (kind) {
        case HEADING:
            valid = detail instanceof HeadingDetail;
            break;

        case NORMAL:
            valid = detail instanceof ParagraphDetail;
            break;

        case BULLET_ITEM:
        case NUMBER_ITEM:
            valid = detail instanceof ListDetail;
            break;

        case CODE_FENCE:
        case CODE_LINE:
        case BLANK:
        case HORIZONTAL_RULE:
        case TABLE_SEPARATOR:
        case TABLE_ROW:
            valid = detail == null;
            break;

        default:
            throw new AssertionError("Unhandled LineKind: " + kind);
        }

        if (!valid) {
            throw new IllegalArgumentException("Detail does not match LineKind: " + kind);
        }
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
        return detail instanceof HeadingDetail ? ((HeadingDetail)detail).level : -1;
    }

    public String getHeadingText() {
        return detail instanceof HeadingDetail ? ((HeadingDetail)detail).text : null;
    }

    public boolean endsWithHardBreak() {
        return detail != null && detail.endsWithHardBreak;
    }

    public String getParagraphText() {
        if (detail instanceof ParagraphDetail) {
            return ((ParagraphDetail)detail).text;
        }

        if (detail instanceof ListDetail) {
            return ((ListDetail)detail).contentText;
        }

        return null;
    }

    public String getListMarkerText() {
        return detail instanceof ListDetail ? ((ListDetail)detail).markerText : null;
    }

    public String getListContentText() {
        return detail instanceof ListDetail ? ((ListDetail)detail).contentText : null;
    }

    public boolean isTableLike() {
        return kind == LineKind.TABLE_SEPARATOR || kind == LineKind.TABLE_ROW;
    }

    private abstract static class Detail {

        final boolean endsWithHardBreak;

        Detail(boolean endsWithHardBreak) {
            this.endsWithHardBreak = endsWithHardBreak;
        }
    }

    private static final class HeadingDetail extends Detail {

        final int level;
        final String text;

        HeadingDetail(int level, String text, boolean endsWithHardBreak) {
            super(endsWithHardBreak);
            this.level = level;
            this.text = text;
        }
    }

    private static final class ParagraphDetail extends Detail {

        final String text;

        ParagraphDetail(String text, boolean endsWithHardBreak) {
            super(endsWithHardBreak);
            this.text = text;
        }
    }

    private static final class ListDetail extends Detail {

        final String markerText;
        final String contentText;

        ListDetail(String markerText, String contentText, boolean endsWithHardBreak) {

            super(endsWithHardBreak);
            this.markerText = markerText;
            this.contentText = contentText;
        }
    }
}