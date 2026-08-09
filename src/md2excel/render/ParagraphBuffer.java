package md2excel.render;

import org.apache.poi.ss.usermodel.CellStyle;

final class ParagraphBuffer {

    // MarkdownInline 側の paragraph parser でも同じ値を使う前提
    static final char SOFT_BREAK_TOKEN = '\uE000';
    static final char HARD_BREAK_TOKEN = '\uE001';

    enum Kind {
        NORMAL,
        QUOTE_NORMAL,
        BULLET,
        NUMBER,
        QUOTE_BULLET,
        QUOTE_NUMBER
    }

    final Kind kind;
    final StringBuilder inlineText = new StringBuilder(128);

    boolean hasAnyLine = false;
    boolean prevLineHardBreak = false;

    // 出力先
    int firstCol = -1;
    int continuationCol = -1;

    // インデント基準
    int baseIndent = 0;

    // 通常段落でだけ使う
    boolean reuseMarkdownBlankForFirstRow = false;
    boolean isListNote = false;

    // 引用段落で使う
    boolean inBlockQuote = false;
    int quoteStartCol = -1; // 本文列
    int quoteDecorCol = -1; // 左罫線列

    // スタイル
    CellStyle firstLineStyle;
    CellStyle continuationStyle;

    // 1行目だけ前置するプレフィックス
    String firstLinePrefix = "";

    ParagraphBuffer(Kind kind) {
        this.kind = kind;
    }

    void appendLine(String lineText, boolean endsWithHardBreak) {
        if (lineText == null) {
            lineText = "";
        }

        if (hasAnyLine) {
            inlineText.append(prevLineHardBreak ? HARD_BREAK_TOKEN : SOFT_BREAK_TOKEN);
        }

        inlineText.append(lineText);
        hasAnyLine = true;
        prevLineHardBreak = endsWithHardBreak;
    }

    String getParagraphText() {
        return inlineText.toString();
    }

    boolean isEmpty() {
        return !hasAnyLine && (firstLinePrefix == null || firstLinePrefix.isEmpty());
    }

    boolean isQuoteKind() {
        return kind == Kind.QUOTE_NORMAL || kind == Kind.QUOTE_BULLET || kind == Kind.QUOTE_NUMBER;
    }
}