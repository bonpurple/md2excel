package md2excel.config;

import java.util.Objects;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Excelシートの設定。
 */
public final class MdSheetSettings {

    /*
     * 左右に1列ずつ余白を確保し、 その間に最低1列の描画範囲を確保するため、最低3列。
     */
    public static final int MIN_TOTAL_COLUMN_COUNT = 3;

    /*
     * Excel 2007以降のファイル形式上の最大列数。
     *
     * 本アプリケーションの設定上限とは区別する。
     */
    public static final int EXCEL_MAX_TOTAL_COLUMN_COUNT = SpreadsheetVersion.EXCEL2007.getMaxColumns();

    /*
     * 引用、コードブロック、水平線などで描画範囲全体に セルを生成するため、メモリ消費を考慮して アプリケーション上の最大列数を256列に制限する。
     */
    public static final int MAX_TOTAL_COLUMN_COUNT = 256;

    private final int totalColumnCount;
    private final VerticalAlignment verticalAlignment;

    public MdSheetSettings(int totalColumnCount, VerticalAlignment verticalAlignment) {

        this.totalColumnCount = validateTotalColumnCount(totalColumnCount);

        if (verticalAlignment == null) {
            throw new IllegalArgumentException("verticalAlignment must not be null");
        }

        this.verticalAlignment = verticalAlignment;
    }

    /**
     * A列から数えた、シート上で初期化する総列数。
     */
    public int getTotalColumnCount() {
        return totalColumnCount;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    private static int validateTotalColumnCount(int value) {

        if (value < MIN_TOTAL_COLUMN_COUNT || value > MAX_TOTAL_COLUMN_COUNT) {

            throw new IllegalArgumentException("totalColumnCount must be between " + MIN_TOTAL_COLUMN_COUNT + " and "
                    + MAX_TOTAL_COLUMN_COUNT + ": " + value);
        }

        return value;
    }

    @Override
    public boolean equals(Object other) {
        if (this == other) {
            return true;
        }

        if (!(other instanceof MdSheetSettings)) {
            return false;
        }

        MdSheetSettings that = (MdSheetSettings) other;

        return totalColumnCount == that.totalColumnCount && verticalAlignment == that.verticalAlignment;
    }

    @Override
    public int hashCode() {
        return Objects.hash(Integer.valueOf(totalColumnCount), verticalAlignment);
    }

    @Override
    public String toString() {
        return "MdSheetSettings{" + "totalColumnCount=" + totalColumnCount + ", verticalAlignment=" + verticalAlignment
                + '}';
    }
}