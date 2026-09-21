package md2excel.excel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertSame;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class MdStyleCatalogTest {

    @Test
    public void createsMainStyles() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MdStyleCatalog styles = createCatalog(workbook);

            assertNotNull(styles.heading1Style);
            assertNotNull(styles.heading2Style);
            assertNotNull(styles.heading3Style);
            assertNotNull(styles.heading4Style);

            assertNotNull(styles.normalStyle);
            assertNotNull(styles.blankRowStyle);
            assertNotNull(styles.bulletStyle);
            assertNotNull(styles.listStyle);

            assertNotNull(styles.codeBlockStyle);
            assertNotNull(styles.horizontalRuleStyle);

            assertNotNull(styles.tableHeaderStyle);
            assertNotNull(styles.tableBodyStyle);
            assertNotNull(styles.tableBodyLastRowStyle);

            assertNotNull(styles.blockQuoteLeftStyle);
            assertNotNull(styles.blockQuoteBodyStyle);
        }
    }

    @Test
    public void headingStylesUseExpectedFonts() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MdStyleCatalog styles = createCatalog(workbook);

            Font heading1Font = workbook.getFontAt(styles.heading1Style.getFontIndex());

            Font heading4Font = workbook.getFontAt(styles.heading4Style.getFontIndex());

            assertEquals("Meiryo", heading1Font.getFontName());

            assertEquals(16, heading1Font.getFontHeightInPoints());

            assertEquals(true, heading1Font.getBold());

            assertEquals(11, heading4Font.getFontHeightInPoints());

            assertEquals(true, heading4Font.getBold());
        }
    }

    @Test
    public void tableStylesKeepExpectedBorders() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MdStyleCatalog styles = createCatalog(workbook);

            assertEquals(BorderStyle.THIN, styles.tableHeaderStyle.getBorderBottom());

            assertEquals(BorderStyle.HAIR, styles.tableBodyStyle.getBorderBottom());

            assertEquals(BorderStyle.NONE, styles.tableBodyLastRowStyle.getBorderBottom());
        }
    }

    @Test
    public void codeBlockFrameMaskZeroReturnsBaseStyle() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MdStyleCatalog styles = createCatalog(workbook);

            assertSame(styles.codeBlockStyle, styles.codeBlockFrameStyle(CodeBlockFrameMask.NONE));
        }
    }

    @Test
    public void codeBlockFrameMaskAppliesRequestedBorders() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MdStyleCatalog styles = createCatalog(workbook);

            assertEquals(BorderStyle.THIN, styles.codeBlockFrameStyle(CodeBlockFrameMask.TOP).getBorderTop());

            assertEquals(BorderStyle.THIN, styles.codeBlockFrameStyle(CodeBlockFrameMask.BOTTOM).getBorderBottom());

            assertEquals(BorderStyle.THIN, styles.codeBlockFrameStyle(CodeBlockFrameMask.LEFT).getBorderLeft());

            assertEquals(BorderStyle.THIN, styles.codeBlockFrameStyle(CodeBlockFrameMask.RIGHT).getBorderRight());

            assertEquals(BorderStyle.THIN, styles.codeBlockFrameStyle(CodeBlockFrameMask.ALL).getBorderTop());

            assertEquals(BorderStyle.THIN, styles.codeBlockFrameStyle(CodeBlockFrameMask.ALL).getBorderBottom());

            assertEquals(BorderStyle.THIN, styles.codeBlockFrameStyle(CodeBlockFrameMask.ALL).getBorderLeft());

            assertEquals(BorderStyle.THIN, styles.codeBlockFrameStyle(CodeBlockFrameMask.ALL).getBorderRight());
        }
    }

    @Test(expected = IllegalArgumentException.class)
    public void rejectsInvalidCodeBlockFrameMask() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MdStyleCatalog styles = createCatalog(workbook);

            styles.codeBlockFrameStyle(CodeBlockFrameMask.COMBINATION_COUNT);
        }
    }

    @Test(expected = IllegalArgumentException.class)
    public void rejectsNegativeCodeBlockFrameMask() throws Exception {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            MdStyleCatalog styles = createCatalog(workbook);

            styles.codeBlockFrameStyle(-1);
        }
    }

    private static MdStyleCatalog createCatalog(XSSFWorkbook workbook) {

        return new MdStyleCatalog(workbook, "Meiryo", 16, 14, 12, 11, VerticalAlignment.BOTTOM);
    }
}