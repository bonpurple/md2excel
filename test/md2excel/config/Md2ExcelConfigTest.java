package md2excel.config;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.Test;

public class Md2ExcelConfigTest {

    @Test
    public void acceptsValidConfigurationAndTrimsText() {
        Md2ExcelConfig config = new Md2ExcelConfig(" input.md ", " output.xlsx ", 40, " Meiryo ", 16, 14, 12, 11,
                VerticalAlignment.BOTTOM);

        assertEquals("input.md", config.inPath);
        assertEquals("output.xlsx", config.outPath);
        assertEquals("Meiryo", config.fontName);
        assertEquals(40, config.sheetColumnCount);
        assertEquals(16, config.h1Size);
        assertEquals(11, config.normalSize);
    }

    @Test
    public void rejectsTooFewColumns() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                createConfig(Md2ExcelConfig.MIN_SHEET_COLUMN_COUNT - 1, 11);
            }
        });
    }

    @Test
    public void rejectsTooManyColumns() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                createConfig(Md2ExcelConfig.MAX_SHEET_COLUMN_COUNT + 1, 11);
            }
        });
    }

    @Test
    public void rejectsFontSizeBelowMinimum() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                createConfig(40, Md2ExcelConfig.MIN_FONT_SIZE - 1);
            }
        });
    }

    @Test
    public void rejectsFontSizeAboveMaximum() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                createConfig(40, Md2ExcelConfig.MAX_FONT_SIZE + 1);
            }
        });
    }

    @Test
    public void rejectsNullVerticalAlignment() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new Md2ExcelConfig("input.md", "output.xlsx", 40, "Meiryo", 16, 14, 12, 11, null);
            }
        });
    }

    private static Md2ExcelConfig createConfig(int columns, int normalSize) {

        return new Md2ExcelConfig("input.md", "output.xlsx", columns, "Meiryo", 16, 14, 12, normalSize,
                VerticalAlignment.BOTTOM);
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