package md2excel.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

public final class MdStyleCatalog {

    public final CellStyle heading1Style;
    public final CellStyle heading2Style;
    public final CellStyle heading3Style;
    public final CellStyle heading4Style;

    public final CellStyle normalStyle;
    public final CellStyle blankRowStyle;
    public final CellStyle bulletStyle;
    public final CellStyle listStyle;

    public final CellStyle codeBlockStyle;
    public final CellStyle horizontalRuleStyle;

    public final CellStyle tableHeaderStyle;
    public final CellStyle tableBodyStyle;
    public final CellStyle tableBodyLastRowStyle;

    public final CellStyle tableHeaderQuoteStyle;
    public final CellStyle tableBodyQuoteStyle;
    public final CellStyle tableBodyLastRowQuoteStyle;

    public final CellStyle blockQuoteLeftStyle;
    public final CellStyle blockQuoteBodyStyle;
    public final CellStyle blockQuoteBlankLeftStyle;
    public final CellStyle blockQuoteBlankBodyStyle;

    public final CellStyle blockQuoteHeading1Style;
    public final CellStyle blockQuoteHeading2Style;
    public final CellStyle blockQuoteHeading3Style;
    public final CellStyle blockQuoteHeading4Style;

    public final CellStyle blockQuoteHorizontalRuleBodyStyle;

    /*
     * インデックスはCodeBlockFrameMaskの組み合わせ。 NONEはcodeBlockStyleを直接使用する。
     */
    private final CellStyle[] codeBlockFrameStyles;

    public MdStyleCatalog(XSSFWorkbook workbook, String fontName, int h1Size, int h2Size, int h3Size, int normalSize,
            VerticalAlignment verticalAlignment) {

        MdFontFactory fontFactory = new MdFontFactory(workbook);

        MdCellStyleFactory styleFactory = new MdCellStyleFactory(workbook);

        XSSFCellStyle baseStyle = styleFactory.createBase(verticalAlignment);

        HeadingStyles headings = createHeadingStyles(baseStyle, fontFactory, styleFactory, fontName, h1Size, h2Size,
                h3Size, normalSize);

        heading1Style = headings.heading1;
        heading2Style = headings.heading2;
        heading3Style = headings.heading3;
        heading4Style = headings.heading4;

        BasicStyles basic = createBasicStyles(baseStyle, fontFactory, styleFactory, fontName, normalSize);

        normalStyle = basic.normal;
        blankRowStyle = basic.blankRow;
        bulletStyle = basic.bullet;
        listStyle = basic.list;

        XSSFColor codeBackground = styleFactory.color(MdStyleDefaults.CODE_BACKGROUND);

        CodeStyles code = createCodeStyles(workbook, baseStyle, fontFactory, styleFactory, codeBackground);

        codeBlockStyle = code.block;
        codeBlockFrameStyles = code.frameStyles;

        horizontalRuleStyle = createHorizontalRuleStyle(blankRowStyle, styleFactory);

        TableStyles tables = createTableStyles(baseStyle, fontFactory, styleFactory, fontName, normalSize);

        tableHeaderStyle = tables.header;
        tableBodyStyle = tables.body;
        tableBodyLastRowStyle = tables.bodyLastRow;

        QuoteStyles quotes = createQuoteStyles(headings, basic, tables, styleFactory, codeBackground);

        tableHeaderQuoteStyle = quotes.tableHeader;

        tableBodyQuoteStyle = quotes.tableBody;

        tableBodyLastRowQuoteStyle = quotes.tableBodyLastRow;

        blockQuoteLeftStyle = quotes.left;

        blockQuoteBodyStyle = quotes.body;

        blockQuoteBlankLeftStyle = quotes.blankLeft;

        blockQuoteBlankBodyStyle = quotes.blankBody;

        blockQuoteHeading1Style = quotes.heading1;

        blockQuoteHeading2Style = quotes.heading2;

        blockQuoteHeading3Style = quotes.heading3;

        blockQuoteHeading4Style = quotes.heading4;

        blockQuoteHorizontalRuleBodyStyle = quotes.horizontalRuleBody;
    }

    private static HeadingStyles createHeadingStyles(CellStyle baseStyle, MdFontFactory fontFactory,
            MdCellStyleFactory styleFactory, String fontName, int h1Size, int h2Size, int h3Size, int normalSize) {

        CellStyle heading1 = styleFactory.cloneWithFont(baseStyle, fontFactory.create(fontName, h1Size, true, false));

        CellStyle heading2 = styleFactory.cloneWithFont(baseStyle, fontFactory.create(fontName, h2Size, true, false));

        CellStyle heading3 = styleFactory.cloneWithFont(baseStyle, fontFactory.create(fontName, h3Size, true, false));

        CellStyle heading4 = styleFactory.cloneWithFont(baseStyle,
                fontFactory.create(fontName, normalSize, true, false));

        return new HeadingStyles(heading1, heading2, heading3, heading4);
    }

    private static BasicStyles createBasicStyles(CellStyle baseStyle, MdFontFactory fontFactory,
            MdCellStyleFactory styleFactory, String fontName, int normalSize) {

        CellStyle normal = styleFactory.cloneWithFont(baseStyle,
                fontFactory.create(fontName, normalSize, false, false));

        CellStyle blankRow = styleFactory.cloneWithFont(baseStyle, fontFactory.create(fontName, 6, false, false));

        CellStyle bullet = styleFactory.cloneOf(normal);

        CellStyle list = styleFactory.cloneOf(normal);

        return new BasicStyles(normal, blankRow, bullet, list);
    }

    private static CodeStyles createCodeStyles(XSSFWorkbook workbook, CellStyle baseStyle, MdFontFactory fontFactory,
            MdCellStyleFactory styleFactory, XSSFColor codeBackground) {

        XSSFCellStyle codeBlock = styleFactory.cloneWithFont(baseStyle,
                fontFactory.create(MdStyleDefaults.CODE_CJK_FONT_NAME, 10, false, false));

        styleFactory.applyFill(codeBlock, codeBackground);

        CellStyle[] frameStyles = createCodeBlockFrameStyles(workbook, codeBlock);

        return new CodeStyles(codeBlock, frameStyles);
    }

    private static CellStyle createHorizontalRuleStyle(CellStyle blankRowStyle, MdCellStyleFactory styleFactory) {

        XSSFCellStyle horizontalRule = styleFactory.cloneOf(blankRowStyle);

        horizontalRule.setBorderBottom(BorderStyle.HAIR);

        return horizontalRule;
    }

    private static TableStyles createTableStyles(CellStyle baseStyle, MdFontFactory fontFactory,
            MdCellStyleFactory styleFactory, String fontName, int normalSize) {

        XSSFCellStyle header = styleFactory.cloneWithFont(baseStyle,
                fontFactory.create(fontName, normalSize, true, false));

        header.setBorderBottom(BorderStyle.THIN);

        XSSFCellStyle body = styleFactory.cloneWithFont(baseStyle,
                fontFactory.create(fontName, normalSize, false, false));

        body.setBorderBottom(BorderStyle.HAIR);

        XSSFCellStyle bodyLastRow = styleFactory.cloneOf(body);

        bodyLastRow.setBorderBottom(BorderStyle.NONE);

        return new TableStyles(header, body, bodyLastRow);
    }

    private static QuoteStyles createQuoteStyles(HeadingStyles headings, BasicStyles basic, TableStyles tables,
            MdCellStyleFactory styleFactory, XSSFColor codeBackground) {

        /*
         * 引用内見出し
         */
        CellStyle heading1 = styleFactory.cloneWithFill(headings.heading1, codeBackground);

        CellStyle heading2 = styleFactory.cloneWithFill(headings.heading2, codeBackground);

        CellStyle heading3 = styleFactory.cloneWithFill(headings.heading3, codeBackground);

        CellStyle heading4 = styleFactory.cloneWithFill(headings.heading4, codeBackground);

        /*
         * 引用内テーブル
         */
        CellStyle tableHeader = styleFactory.cloneWithFill(tables.header, codeBackground);

        CellStyle tableBody = styleFactory.cloneWithFill(tables.body, codeBackground);

        CellStyle tableBodyLastRow = styleFactory.cloneWithFill(tables.bodyLastRow, codeBackground);

        /*
         * 引用本文
         */
        XSSFCellStyle body = styleFactory.cloneWithFill(basic.normal, codeBackground);

        styleFactory.clearBorders(body);

        XSSFColor quoteBorder = styleFactory.color(MdStyleDefaults.QUOTE_BORDER);

        XSSFCellStyle left = styleFactory.cloneOf(body);

        left.setBorderLeft(BorderStyle.THICK);

        left.setBorderColor(BorderSide.LEFT, quoteBorder);

        /*
         * 引用空行
         */
        XSSFCellStyle blankBody = styleFactory.cloneWithFill(basic.blankRow, codeBackground);

        styleFactory.clearBorders(blankBody);

        XSSFCellStyle blankLeft = styleFactory.cloneOf(blankBody);

        blankLeft.setBorderLeft(BorderStyle.THICK);

        blankLeft.setBorderColor(BorderSide.LEFT, quoteBorder);

        /*
         * 引用内水平線
         */
        XSSFCellStyle horizontalRuleBody = styleFactory.cloneOf(blankBody);

        horizontalRuleBody.setBorderBottom(BorderStyle.HAIR);

        return new QuoteStyles(tableHeader, tableBody, tableBodyLastRow, left, body, blankLeft, blankBody, heading1,
                heading2, heading3, heading4, horizontalRuleBody);
    }

    private static CellStyle[] createCodeBlockFrameStyles(XSSFWorkbook workbook, CellStyle codeBlockStyle) {

        CellStyle[] frameStyles = new CellStyle[CodeBlockFrameMask.COMBINATION_COUNT];

        /*
         * mask == 0の場合はcodeBlockStyleを直接返すため、 frameStyles[0]は使用しない。
         */
        for (int mask = CodeBlockFrameMask.TOP; mask < frameStyles.length; mask++) {

            XSSFCellStyle style = workbook.createCellStyle();

            style.cloneStyleFrom(codeBlockStyle);

            if (CodeBlockFrameMask.contains(mask, CodeBlockFrameMask.TOP)) {

                style.setBorderTop(BorderStyle.THIN);
            }

            if (CodeBlockFrameMask.contains(mask, CodeBlockFrameMask.BOTTOM)) {

                style.setBorderBottom(BorderStyle.THIN);
            }

            if (CodeBlockFrameMask.contains(mask, CodeBlockFrameMask.LEFT)) {

                style.setBorderLeft(BorderStyle.THIN);
            }

            if (CodeBlockFrameMask.contains(mask, CodeBlockFrameMask.RIGHT)) {

                style.setBorderRight(BorderStyle.THIN);
            }

            frameStyles[mask] = style;
        }

        return frameStyles;
    }

    public CellStyle codeBlockFrameStyle(int mask) {
        if (!CodeBlockFrameMask.isValid(mask)) {
            throw new IllegalArgumentException("Invalid code block frame mask: " + mask);
        }

        if (mask == CodeBlockFrameMask.NONE) {
            return codeBlockStyle;
        }

        return codeBlockFrameStyles[mask];
    }

    private static final class HeadingStyles {

        final CellStyle heading1;
        final CellStyle heading2;
        final CellStyle heading3;
        final CellStyle heading4;

        HeadingStyles(CellStyle heading1, CellStyle heading2, CellStyle heading3, CellStyle heading4) {

            this.heading1 = heading1;
            this.heading2 = heading2;
            this.heading3 = heading3;
            this.heading4 = heading4;
        }
    }

    private static final class BasicStyles {

        final CellStyle normal;
        final CellStyle blankRow;
        final CellStyle bullet;
        final CellStyle list;

        BasicStyles(CellStyle normal, CellStyle blankRow, CellStyle bullet, CellStyle list) {

            this.normal = normal;
            this.blankRow = blankRow;
            this.bullet = bullet;
            this.list = list;
        }
    }

    private static final class CodeStyles {

        final CellStyle block;
        final CellStyle[] frameStyles;

        CodeStyles(CellStyle block, CellStyle[] frameStyles) {

            this.block = block;
            this.frameStyles = frameStyles;
        }
    }

    private static final class TableStyles {

        final CellStyle header;
        final CellStyle body;
        final CellStyle bodyLastRow;

        TableStyles(CellStyle header, CellStyle body, CellStyle bodyLastRow) {

            this.header = header;
            this.body = body;
            this.bodyLastRow = bodyLastRow;
        }
    }

    private static final class QuoteStyles {

        final CellStyle tableHeader;
        final CellStyle tableBody;
        final CellStyle tableBodyLastRow;

        final CellStyle left;
        final CellStyle body;
        final CellStyle blankLeft;
        final CellStyle blankBody;

        final CellStyle heading1;
        final CellStyle heading2;
        final CellStyle heading3;
        final CellStyle heading4;

        final CellStyle horizontalRuleBody;

        QuoteStyles(CellStyle tableHeader, CellStyle tableBody, CellStyle tableBodyLastRow, CellStyle left,
                CellStyle body, CellStyle blankLeft, CellStyle blankBody, CellStyle heading1, CellStyle heading2,
                CellStyle heading3, CellStyle heading4, CellStyle horizontalRuleBody) {

            this.tableHeader = tableHeader;
            this.tableBody = tableBody;
            this.tableBodyLastRow = tableBodyLastRow;

            this.left = left;
            this.body = body;
            this.blankLeft = blankLeft;
            this.blankBody = blankBody;

            this.heading1 = heading1;
            this.heading2 = heading2;
            this.heading3 = heading3;
            this.heading4 = heading4;

            this.horizontalRuleBody = horizontalRuleBody;
        }
    }
}