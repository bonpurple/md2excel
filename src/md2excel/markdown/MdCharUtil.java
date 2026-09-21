package md2excel.markdown;

/**
 * Markdown字句解析で使用する文字判定。
 */
public final class MdCharUtil {

    private MdCharUtil() {
    }

    /**
     * CommonMarkのバックスラッシュエスケープ対象となる ASCII punctuationかどうかを判定する。
     */
    public static boolean isAsciiPunctuation(char ch) {
        if (ch > 0x7F) {
            return false;
        }

        return (ch >= '!' && ch <= '/') || (ch >= ':' && ch <= '@') || (ch >= '[' && ch <= '`')
                || (ch >= '{' && ch <= '~');
    }
}