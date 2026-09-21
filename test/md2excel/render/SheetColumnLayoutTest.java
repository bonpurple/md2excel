package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import org.junit.Test;

public class SheetColumnLayoutTest {

    @Test
    public void fortyColumnsWithOuterMarginsRendersFromBThroughAM() {
        SheetColumnLayout layout = new SheetColumnLayout(40, 1, 1);

        assertEquals(40, layout.getTotalColumnCount());

        assertEquals(1, layout.getLeftMarginColumnCount());

        assertEquals(1, layout.getRightMarginColumnCount());

        // B列
        assertEquals(1, layout.getRenderStartColIndex());

        // 描画終了列の次はAN列
        assertEquals(39, layout.getRenderEndColExclusive());

        // 最終描画列はAM列
        assertEquals(38, layout.getRenderLastColIndex());

        assertEquals(38, layout.getRenderableColumnCount());

        assertTrue(layout.isRenderableColumn(1));

        assertTrue(layout.isRenderableColumn(38));

        assertFalse(layout.isRenderableColumn(0));

        assertFalse(layout.isRenderableColumn(39));
    }

    @Test
    public void minimumThreeColumnsLeavesOnlyBRenderable() {
        SheetColumnLayout layout = new SheetColumnLayout(3, 1, 1);

        assertEquals(1, layout.getRenderStartColIndex());

        assertEquals(2, layout.getRenderEndColExclusive());

        assertEquals(1, layout.getRenderLastColIndex());

        assertEquals(1, layout.getRenderableColumnCount());

        assertTrue(layout.isRenderableColumn(1));
    }

    @Test
    public void supportsLayoutWithoutMargins() {
        SheetColumnLayout layout = new SheetColumnLayout(3, 0, 0);

        assertEquals(0, layout.getRenderStartColIndex());

        assertEquals(3, layout.getRenderEndColExclusive());

        assertEquals(2, layout.getRenderLastColIndex());

        assertEquals(3, layout.getRenderableColumnCount());
    }

    @Test
    public void rejectsLayoutWithoutRenderableColumns() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new SheetColumnLayout(2, 1, 1);
            }
        });
    }

    @Test
    public void rejectsNegativeLeftMargin() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new SheetColumnLayout(40, -1, 1);
            }
        });
    }

    @Test
    public void rejectsNegativeRightMargin() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new SheetColumnLayout(40, 1, -1);
            }
        });
    }

    @Test
    public void acceptsExcelPhysicalMaximum() {
        SheetColumnLayout layout = new SheetColumnLayout(SheetColumnLayout.EXCEL_MAX_TOTAL_COLUMN_COUNT, 1, 1);

        assertEquals(SheetColumnLayout.EXCEL_MAX_TOTAL_COLUMN_COUNT, layout.getTotalColumnCount());
    }

    private static void assertInvalid(Runnable action) {
        try {
            action.run();
            fail("IllegalArgumentException was expected");
        } catch (IllegalArgumentException expected) {
            // expected
        }
    }
}