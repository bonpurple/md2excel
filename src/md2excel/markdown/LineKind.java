package md2excel.markdown;

/**
 * Markdown行の解析結果を表す種類。
 *
 * レンダリング固有の処理には依存しない。
 */
public enum LineKind {
    CODE_FENCE,
    CODE_LINE,
    BLANK,
    HORIZONTAL_RULE,
    TABLE_SEPARATOR,
    TABLE_ROW,
    HEADING,
    BULLET_ITEM,
    NUMBER_ITEM,
    NORMAL
}