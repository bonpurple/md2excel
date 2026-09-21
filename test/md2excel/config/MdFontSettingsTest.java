package md2excel.config;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import org.junit.Test;

public class MdFontSettingsTest {

    @Test
    public void acceptsValidSettingsAndTrimsFontName() {
        MdFontSettings settings = new MdFontSettings(" Meiryo ", 16, 14, 12, 11);

        assertEquals("Meiryo", settings.getFontName());

        assertEquals(16, settings.getH1Size());
        assertEquals(14, settings.getH2Size());
        assertEquals(12, settings.getH3Size());
        assertEquals(11, settings.getNormalSize());
    }

    @Test
    public void rejectsNullFontName() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new MdFontSettings(null, 16, 14, 12, 11);
            }
        });
    }

    @Test
    public void rejectsEmptyFontName() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new MdFontSettings("   ", 16, 14, 12, 11);
            }
        });
    }

    @Test
    public void rejectsFontSizeBelowMinimum() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new MdFontSettings("Meiryo", 16, 14, 12, MdFontSettings.MIN_FONT_SIZE - 1);
            }
        });
    }

    @Test
    public void rejectsFontSizeAboveMaximum() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new MdFontSettings("Meiryo", MdFontSettings.MAX_FONT_SIZE + 1, 14, 12, 11);
            }
        });
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