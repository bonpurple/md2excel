package md2excel.render;

/**
 * 段落内で入力行の改行種別を保持するための内部トークン。
 */
final class ParagraphBreakTokens {

    static final char SOFT_BREAK_TOKEN = '\uE000';
    static final char HARD_BREAK_TOKEN = '\uE001';

    private ParagraphBreakTokens() {
    }
}