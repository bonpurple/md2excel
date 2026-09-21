package md2excel.config;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertSame;
import static org.junit.Assert.fail;

import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.Test;

public class Md2ExcelConfigTest {

    @Test
    public void acceptsValidConfiguration() {
        Path inputPath = Paths.get("input.md");

        Path outputPath = Paths.get("output.xlsx");

        MdFontSettings fontSettings = new MdFontSettings("Meiryo", 16, 14, 12, 11);

        MdSheetSettings sheetSettings = new MdSheetSettings(40, VerticalAlignment.BOTTOM);

        Md2ExcelConfig config = new Md2ExcelConfig(inputPath, outputPath, fontSettings, sheetSettings);

        assertEquals(inputPath, config.getInputPath());

        assertEquals(outputPath, config.getOutputPath());

        assertSame(fontSettings, config.getFontSettings());

        assertSame(sheetSettings, config.getSheetSettings());
    }

    @Test
    public void rejectsNullInputPath() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new Md2ExcelConfig(null, Paths.get("output.xlsx"), validFontSettings(), validSheetSettings());
            }
        });
    }

    @Test
    public void rejectsEmptyInputPath() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new Md2ExcelConfig(Paths.get(""), Paths.get("output.xlsx"), validFontSettings(), validSheetSettings());
            }
        });
    }

    @Test
    public void rejectsNullOutputPath() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new Md2ExcelConfig(Paths.get("input.md"), null, validFontSettings(), validSheetSettings());
            }
        });
    }

    @Test
    public void rejectsEmptyOutputPath() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new Md2ExcelConfig(Paths.get("input.md"), Paths.get(""), validFontSettings(), validSheetSettings());
            }
        });
    }

    @Test
    public void rejectsNullFontSettings() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new Md2ExcelConfig(Paths.get("input.md"), Paths.get("output.xlsx"), null, validSheetSettings());
            }
        });
    }

    @Test
    public void rejectsNullSheetSettings() {
        assertInvalid(new Runnable() {
            @Override
            public void run() {
                new Md2ExcelConfig(Paths.get("input.md"), Paths.get("output.xlsx"), validFontSettings(), null);
            }
        });
    }

    private static MdFontSettings validFontSettings() {
        return new MdFontSettings("Meiryo", 16, 14, 12, 11);
    }

    private static MdSheetSettings validSheetSettings() {
        return new MdSheetSettings(40, VerticalAlignment.BOTTOM);
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