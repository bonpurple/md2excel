package md2excel.excel;

/**
 * コードブロック外枠の適用位置を表すビットマスク。
 */
public final class CodeBlockFrameMask {

    public static final int NONE = 0;

    public static final int TOP = 1 << 0;
    public static final int BOTTOM = 1 << 1;
    public static final int LEFT = 1 << 2;
    public static final int RIGHT = 1 << 3;

    /**
     * TOP、BOTTOM、LEFT、RIGHTの全組み合わせ数。
     *
     * 0～15を有効なマスクとして使用する。
     */
    public static final int COMBINATION_COUNT = 1 << 4;

    public static final int ALL = TOP | BOTTOM | LEFT | RIGHT;

    private CodeBlockFrameMask() {
    }

    /**
     * 指定した辺がマスクに含まれるか判定する。
     */
    public static boolean contains(int mask, int side) {

        return (mask & side) != 0;
    }

    /**
     * 有効なコード枠マスクか判定する。
     */
    public static boolean isValid(int mask) {
        return mask >= NONE && mask < COMBINATION_COUNT;
    }
}