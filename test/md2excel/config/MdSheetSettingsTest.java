package md2excel.config;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.Test;

public class MdSheetSettingsTest {

    @Test
    public void acceptsValidSettings() {
        MdSheetSettings settings = new MdSheetSettings(40, VerticalAlignment.BOTTOM);

        assertEquals(40, settings.getTotalColumnCount());

        assertEquals(VerticalAlignment.BOTTOM, settings.getVerticalAlignment());
    }

    @Test
    public void rejectsTooFewColumns() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new MdSheetSettings(MdSheetSettings.MIN_TOTAL_COLUMN_COUNT - 1, VerticalAlignment.BOTTOM);
            }
        });
    }

    @Test
    public void rejectsTooManyColumns() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new MdSheetSettings(MdSheetSettings.MAX_TOTAL_COLUMN_COUNT + 1, VerticalAlignment.BOTTOM);
            }
        });
    }

    @Test
    public void rejectsNullVerticalAlignment() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new MdSheetSettings(40, null);
            }
        });
    }

    @Test
    public void acceptsApplicationMaximumColumns() {
        MdSheetSettings settings = new MdSheetSettings(MdSheetSettings.MAX_TOTAL_COLUMN_COUNT,
                VerticalAlignment.BOTTOM);

        assertEquals(256, settings.getTotalColumnCount());
    }

    @Test
    public void applicationMaximumDoesNotExceedExcelMaximum() {
        assertTrue(MdSheetSettings.MAX_TOTAL_COLUMN_COUNT <= MdSheetSettings.EXCEL_MAX_TOTAL_COLUMN_COUNT);
    }

    @Test
    public void acceptsColumnCountBoundaries() {
        MdSheetSettings minimum = new MdSheetSettings(MdSheetSettings.MIN_TOTAL_COLUMN_COUNT, VerticalAlignment.BOTTOM);

        MdSheetSettings maximum = new MdSheetSettings(MdSheetSettings.MAX_TOTAL_COLUMN_COUNT, VerticalAlignment.BOTTOM);

        assertEquals(MdSheetSettings.MIN_TOTAL_COLUMN_COUNT, minimum.getTotalColumnCount());

        assertEquals(MdSheetSettings.MAX_TOTAL_COLUMN_COUNT, maximum.getTotalColumnCount());
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