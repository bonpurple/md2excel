package md2excel.config;

import static org.junit.Assert.assertEquals;

import java.nio.file.Path;
import java.nio.file.Paths;

import org.junit.Test;

public class Md2ExcelConfigDialogTest {

    @Test
    public void replacesExistingExtension() {
        Path input = Paths.get("work", "document.md");

        Path output = Md2ExcelConfigDialog.replaceExtension(input, ".xlsx");

        assertEquals(Paths.get("work", "document.xlsx"), output);
    }

    @Test
    public void appendsExtensionWhenMissing() {
        Path input = Paths.get("work", "document");

        Path output = Md2ExcelConfigDialog.replaceExtension(input, ".xlsx");

        assertEquals(Paths.get("work", "document.xlsx"), output);
    }

    @Test
    public void keepsLeadingDotFileName() {
        Path input = Paths.get("work", ".markdown");

        Path output = Md2ExcelConfigDialog.replaceExtension(input, ".xlsx");

        assertEquals(Paths.get("work", ".markdown.xlsx"), output);
    }

    @Test
    public void replacesOnlyLastExtension() {
        Path input = Paths.get("work", "document.spec.md");

        Path output = Md2ExcelConfigDialog.replaceExtension(input, ".xlsx");

        assertEquals(Paths.get("work", "document.spec.xlsx"), output);
    }
}