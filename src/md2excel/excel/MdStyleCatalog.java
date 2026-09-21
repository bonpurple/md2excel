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

    // mask bit: 1=TOP, 2=BOTTOM, 4=LEFT, 8=RIGHT
    private final CellStyle[] codeBlockFrameStyles = new CellStyle[16];

    public MdStyleCatalog(XSSFWorkbook workbook, String fontName, int h1Size, int h2Size, int h3Size, int normalSize,
            VerticalAlignment verticalAlignment) {

        MdFontFactory fonts = new MdFontFactory(workbook);

        MdCellStyleFactory styles = new MdCellStyleFactory(workbook);

        XSSFCellStyle base = styles.createBase(verticalAlignment);

        // 見出し
        heading1Style = styles.cloneWithFont(base, fonts.create(fontName, h1Size, true, false));

        heading2Style = styles.cloneWithFont(base, fonts.create(fontName, h2Size, true, false));

        heading3Style = styles.cloneWithFont(base, fonts.create(fontName, h3Size, true, false));

        heading4Style = styles.cloneWithFont(base, fonts.create(fontName, normalSize, true, false));

        // 通常・空行・リスト
        normalStyle = styles.cloneWithFont(base, fonts.create(fontName, normalSize, false, false));

        blankRowStyle = styles.cloneWithFont(base, fonts.create(fontName, 6, false, false));

        bulletStyle = styles.cloneOf(normalStyle);
        listStyle = styles.cloneOf(normalStyle);

        XSSFColor codeBackground = styles.color(MdStyleDefaults.CODE_BACKGROUND);

        // コードブロック
        XSSFCellStyle codeStyle = styles.cloneWithFont(base,
                fonts.create(MdStyleDefaults.CODE_CJK_FONT_NAME, 10, false, false));

        styles.applyFill(codeStyle, codeBackground);
        codeBlockStyle = codeStyle;

        initCodeBlockFrameStyles(workbook);

        // 水平線
        XSSFCellStyle horizontalRule = styles.cloneOf(blankRowStyle);

        horizontalRule.setBorderBottom(BorderStyle.HAIR);
        horizontalRuleStyle = horizontalRule;

        // テーブルヘッダー
        XSSFCellStyle tableHeader = styles.cloneWithFont(base, fonts.create(fontName, normalSize, true, false));

        tableHeader.setBorderBottom(BorderStyle.THIN);
        tableHeaderStyle = tableHeader;

        // テーブル本文
        XSSFCellStyle tableBody = styles.cloneWithFont(base, fonts.create(fontName, normalSize, false, false));

        tableBody.setBorderBottom(BorderStyle.HAIR);
        tableBodyStyle = tableBody;

        XSSFCellStyle tableBodyLast = styles.cloneOf(tableBodyStyle);

        tableBodyLast.setBorderBottom(BorderStyle.NONE);
        tableBodyLastRowStyle = tableBodyLast;

        // 引用内見出し
        blockQuoteHeading1Style = styles.cloneWithFill(heading1Style, codeBackground);

        blockQuoteHeading2Style = styles.cloneWithFill(heading2Style, codeBackground);

        blockQuoteHeading3Style = styles.cloneWithFill(heading3Style, codeBackground);

        blockQuoteHeading4Style = styles.cloneWithFill(heading4Style, codeBackground);

        // 引用内テーブル
        tableHeaderQuoteStyle = styles.cloneWithFill(tableHeaderStyle, codeBackground);

        tableBodyQuoteStyle = styles.cloneWithFill(tableBodyStyle, codeBackground);

        tableBodyLastRowQuoteStyle = styles.cloneWithFill(tableBodyLastRowStyle, codeBackground);

        // 引用本文
        XSSFCellStyle quoteBody = styles.cloneWithFill(normalStyle, codeBackground);

        styles.clearBorders(quoteBody);
        blockQuoteBodyStyle = quoteBody;

        XSSFColor quoteBorder = styles.color(MdStyleDefaults.QUOTE_BORDER);

        XSSFCellStyle quoteLeft = styles.cloneOf(blockQuoteBodyStyle);

        quoteLeft.setBorderLeft(BorderStyle.THICK);
        quoteLeft.setBorderColor(BorderSide.LEFT, quoteBorder);

        blockQuoteLeftStyle = quoteLeft;

        // 引用空行
        XSSFCellStyle quoteBlankBody = styles.cloneWithFill(blankRowStyle, codeBackground);

        styles.clearBorders(quoteBlankBody);
        blockQuoteBlankBodyStyle = quoteBlankBody;

        XSSFCellStyle quoteBlankLeft = styles.cloneOf(blockQuoteBlankBodyStyle);

        quoteBlankLeft.setBorderLeft(BorderStyle.THICK);
        quoteBlankLeft.setBorderColor(BorderSide.LEFT, quoteBorder);

        blockQuoteBlankLeftStyle = quoteBlankLeft;

        // 引用内水平線
        XSSFCellStyle quoteHorizontalRule = styles.cloneOf(blockQuoteBlankBodyStyle);

        quoteHorizontalRule.setBorderBottom(BorderStyle.HAIR);
        blockQuoteHorizontalRuleBodyStyle = quoteHorizontalRule;
    }

    private void initCodeBlockFrameStyles(XSSFWorkbook workbook) {

        for (int mask = 1; mask < codeBlockFrameStyles.length; mask++) {
            XSSFCellStyle style = workbook.createCellStyle();
            style.cloneStyleFrom(codeBlockStyle);

            if ((mask & 1) != 0) {
                style.setBorderTop(BorderStyle.THIN);
            }

            if ((mask & 2) != 0) {
                style.setBorderBottom(BorderStyle.THIN);
            }

            if ((mask & 4) != 0) {
                style.setBorderLeft(BorderStyle.THIN);
            }

            if ((mask & 8) != 0) {
                style.setBorderRight(BorderStyle.THIN);
            }

            codeBlockFrameStyles[mask] = style;
        }
    }

    public CellStyle codeBlockFrameStyle(int mask) {
        if (mask < 0 || mask >= codeBlockFrameStyles.length) {
            throw new IllegalArgumentException("Invalid code block frame mask: " + mask);
        }

        if (mask == 0) {
            return codeBlockStyle;
        }

        return codeBlockFrameStyles[mask];
    }
}