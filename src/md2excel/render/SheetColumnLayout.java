package md2excel.render;

import java.util.Objects;

import org.apache.poi.ss.SpreadsheetVersion;

/**
 * Excelシート上の列範囲を表す値オブジェクト。
 *
 * 列番号はApache POIと同じ0始まり。
 */
public final class SheetColumnLayout {

    /**
     * XLSX形式で使用できる物理的な最大列数。
     *
     * アプリケーションの設定上限は MdSheetSettings.MAX_TOTAL_COLUMN_COUNTで定義する。
     */
    public static final int EXCEL_MAX_TOTAL_COLUMN_COUNT = SpreadsheetVersion.EXCEL2007.getMaxColumns();

    private final int totalColumnCount;
    private final int leftMarginColumnCount;
    private final int rightMarginColumnCount;

    public SheetColumnLayout(int totalColumnCount, int leftMarginColumnCount, int rightMarginColumnCount) {

        if (totalColumnCount <= 0 || totalColumnCount > EXCEL_MAX_TOTAL_COLUMN_COUNT) {

            throw new IllegalArgumentException(
                    "totalColumnCount must be between 1 and " + EXCEL_MAX_TOTAL_COLUMN_COUNT + ": " + totalColumnCount);
        }

        if (leftMarginColumnCount < 0) {
            throw new IllegalArgumentException("leftMarginColumnCount must not be negative: " + leftMarginColumnCount);
        }

        if (rightMarginColumnCount < 0) {
            throw new IllegalArgumentException(
                    "rightMarginColumnCount must not be negative: " + rightMarginColumnCount);
        }

        int renderableColumnCount = totalColumnCount - leftMarginColumnCount - rightMarginColumnCount;

        if (renderableColumnCount <= 0) {
            throw new IllegalArgumentException("At least one renderable column is required: " + "totalColumnCount="
                    + totalColumnCount + ", leftMarginColumnCount=" + leftMarginColumnCount
                    + ", rightMarginColumnCount=" + rightMarginColumnCount);
        }

        this.totalColumnCount = totalColumnCount;
        this.leftMarginColumnCount = leftMarginColumnCount;
        this.rightMarginColumnCount = rightMarginColumnCount;
    }

    /**
     * シート上で初期化する総列数。
     */
    public int getTotalColumnCount() {
        return totalColumnCount;
    }

    /**
     * 左余白として確保する列数。
     */
    public int getLeftMarginColumnCount() {
        return leftMarginColumnCount;
    }

    /**
     * 右余白として確保する列数。
     */
    public int getRightMarginColumnCount() {
        return rightMarginColumnCount;
    }

    /**
     * 描画開始列。範囲に含まれる。
     */
    public int getRenderStartColIndex() {
        return leftMarginColumnCount;
    }

    /**
     * 描画終了列の次の列。範囲には含まれない。
     */
    public int getRenderEndColExclusive() {
        return totalColumnCount - rightMarginColumnCount;
    }

    /**
     * 描画範囲内の最後の列。
     */
    public int getRenderLastColIndex() {
        return getRenderEndColExclusive() - 1;
    }

    /**
     * 実際に描画へ使用できる列数。
     */
    public int getRenderableColumnCount() {
        return getRenderEndColExclusive() - getRenderStartColIndex();
    }

    public boolean isRenderableColumn(int columnIndex) {
        return columnIndex >= getRenderStartColIndex() && columnIndex < getRenderEndColExclusive();
    }

    @Override
    public boolean equals(Object other) {
        if (this == other) {
            return true;
        }

        if (!(other instanceof SheetColumnLayout)) {
            return false;
        }

        SheetColumnLayout that = (SheetColumnLayout) other;

        return totalColumnCount == that.totalColumnCount && leftMarginColumnCount == that.leftMarginColumnCount
                && rightMarginColumnCount == that.rightMarginColumnCount;
    }

    @Override
    public int hashCode() {
        return Objects.hash(Integer.valueOf(totalColumnCount), Integer.valueOf(leftMarginColumnCount),
                Integer.valueOf(rightMarginColumnCount));
    }

    @Override
    public String toString() {
        return "SheetColumnLayout{" + "totalColumnCount=" + totalColumnCount + ", leftMarginColumnCount="
                + leftMarginColumnCount + ", rightMarginColumnCount=" + rightMarginColumnCount
                + ", renderStartColIndex=" + getRenderStartColIndex() + ", renderEndColExclusive="
                + getRenderEndColExclusive() + '}';
    }
}