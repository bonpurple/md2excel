package md2excel.excel;

import java.awt.Color;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

public final class MdStyle {
    public final CellStyle heading1Style;
    public final CellStyle heading2Style;
    public final CellStyle heading3Style;
    public final CellStyle heading4Style;
    public final CellStyle normalStyle;
    public final CellStyle blankRowStyle;
    public final CellStyle bulletStyle;
    public final CellStyle listStyle;

    public final CellStyle codeBlockStyle; // 背景＋フォント（枠線なし）

    public final CellStyle horizontalRuleStyle;

    public final CellStyle tableHeaderStyle;
    public final CellStyle tableBodyStyle;
    public final CellStyle tableBodyLastRowStyle;

    public final CellStyle tableHeaderQuoteStyle;
    public final CellStyle tableBodyQuoteStyle;
    public final CellStyle tableBodyLastRowQuoteStyle;

    // 引用ブロック
    public final CellStyle blockQuoteLeftStyle;
    public final CellStyle blockQuoteBodyStyle;
    public final CellStyle blockQuoteBlankLeftStyle;
    public final CellStyle blockQuoteBlankBodyStyle;
    public final CellStyle blockQuoteHeading1Style;
    public final CellStyle blockQuoteHeading2Style;
    public final CellStyle blockQuoteHeading3Style;
    public final CellStyle blockQuoteHeading4Style;
    public final CellStyle blockQuoteHorizontalRuleBodyStyle;

    // コードブロック枠線スタイル（mask で取り出す）
    // mask bit: 1=TOP, 2=BOTTOM, 4=LEFT, 8=RIGHT
    private final CellStyle[] codeBlockFrameStyles = new CellStyle[16];

    public MdStyle(Workbook wb, String fontName, int h1Size, int h2Size, int h3Size, int normalSize,
            VerticalAlignment vAlign) {

        // 共通ベース
        CellStyle base = wb.createCellStyle();
        base.setWrapText(false);
        base.setVerticalAlignment(vAlign);

        // 見出し1
        heading1Style = wb.createCellStyle();
        heading1Style.cloneStyleFrom(base);
        Font h1Font = wb.createFont();
        h1Font.setBold(true);
        h1Font.setFontHeightInPoints((short) h1Size);
        h1Font.setFontName(fontName);
        heading1Style.setFont(h1Font);

        // 見出し2
        heading2Style = wb.createCellStyle();
        heading2Style.cloneStyleFrom(base);
        Font h2Font = wb.createFont();
        h2Font.setBold(true);
        h2Font.setFontHeightInPoints((short) h2Size);
        h2Font.setFontName(fontName);
        heading2Style.setFont(h2Font);

        // 見出し3
        heading3Style = wb.createCellStyle();
        heading3Style.cloneStyleFrom(base);
        Font h3Font = wb.createFont();
        h3Font.setBold(true);
        h3Font.setFontHeightInPoints((short) h3Size);
        h3Font.setFontName(fontName);
        heading3Style.setFont(h3Font);

        // 見出し4+
        heading4Style = wb.createCellStyle();
        heading4Style.cloneStyleFrom(base);
        Font h4Font = wb.createFont();
        h4Font.setBold(true);
        h4Font.setFontHeightInPoints((short) normalSize);
        h4Font.setFontName(fontName);
        heading4Style.setFont(h4Font);

        // 通常
        normalStyle = wb.createCellStyle();
        normalStyle.cloneStyleFrom(base);
        Font normalFont = wb.createFont();
        normalFont.setFontHeightInPoints((short) normalSize);
        normalFont.setFontName(fontName);
        normalStyle.setFont(normalFont);

        // 空行用
        blankRowStyle = wb.createCellStyle();
        blankRowStyle.cloneStyleFrom(base);
        Font blankFont = wb.createFont();
        blankFont.setFontHeightInPoints((short) 6);
        blankFont.setFontName(fontName);
        blankRowStyle.setFont(blankFont);

        // 箇条書き/番号付き（フォントは同じ）
        bulletStyle = wb.createCellStyle();
        bulletStyle.cloneStyleFrom(normalStyle);

        listStyle = wb.createCellStyle();
        listStyle.cloneStyleFrom(normalStyle);

        // ---------------- コードブロック（ベース） ----------------
        CellStyle baseCode = wb.createCellStyle();
        baseCode.cloneStyleFrom(base);
        Font codeFont = wb.createFont();
        codeFont.setFontName("Meiryo");
        codeFont.setFontHeightInPoints((short) 10);
        baseCode.setFont(codeFont);

        XSSFCellStyle xssfBaseCode = (XSSFCellStyle) baseCode;
        XSSFColor bg = new XSSFColor(new Color(232, 232, 232), null);
        xssfBaseCode.setFillForegroundColor(bg);
        xssfBaseCode.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        this.codeBlockStyle = xssfBaseCode;

        // コード枠線スタイル（最大15個だけ作る）
        initCodeBlockFrameStyles(wb);

        // ---------------- 水平線 ----------------
        horizontalRuleStyle = wb.createCellStyle();
        horizontalRuleStyle.cloneStyleFrom(blankRowStyle);
        horizontalRuleStyle.setBorderBottom(BorderStyle.HAIR);

        // ---------------- テーブル ----------------
        CellStyle th = wb.createCellStyle();
        th.cloneStyleFrom(base);
        Font thFont = wb.createFont();
        thFont.setBold(true);
        thFont.setFontHeightInPoints((short) normalSize);
        thFont.setFontName(fontName);
        th.setFont(thFont);
        th.setBorderBottom(BorderStyle.THIN);
        tableHeaderStyle = th;

        CellStyle tb = wb.createCellStyle();
        tb.cloneStyleFrom(base);
        Font tbFont = wb.createFont();
        tbFont.setFontHeightInPoints((short) normalSize);
        tbFont.setFontName(fontName);
        tb.setFont(tbFont);
        tb.setBorderBottom(BorderStyle.HAIR);
        tableBodyStyle = tb;

        CellStyle tbLast = wb.createCellStyle();
        tbLast.cloneStyleFrom(tableBodyStyle);
        tbLast.setBorderBottom(BorderStyle.NONE);
        tableBodyLastRowStyle = tbLast;

        // 引用ブロック（通常行）
        XSSFColor codeBg = ((XSSFCellStyle) this.codeBlockStyle).getFillForegroundXSSFColor();

        this.blockQuoteHeading1Style = createBlockQuoteContentStyle(wb, this.heading1Style, codeBg);
        this.blockQuoteHeading2Style = createBlockQuoteContentStyle(wb, this.heading2Style, codeBg);
        this.blockQuoteHeading3Style = createBlockQuoteContentStyle(wb, this.heading3Style, codeBg);
        this.blockQuoteHeading4Style = createBlockQuoteContentStyle(wb, this.heading4Style, codeBg);

        XSSFCellStyle tableHeaderQuote = (XSSFCellStyle) wb.createCellStyle();
        tableHeaderQuote.cloneStyleFrom(this.tableHeaderStyle);
        if (codeBg != null) {
            tableHeaderQuote.setFillForegroundColor(codeBg);
            tableHeaderQuote.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        this.tableHeaderQuoteStyle = tableHeaderQuote;

        XSSFCellStyle tableBodyQuote = (XSSFCellStyle) wb.createCellStyle();
        tableBodyQuote.cloneStyleFrom(this.tableBodyStyle);
        if (codeBg != null) {
            tableBodyQuote.setFillForegroundColor(codeBg);
            tableBodyQuote.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        this.tableBodyQuoteStyle = tableBodyQuote;

        XSSFCellStyle tableBodyLastRowQuote = (XSSFCellStyle) wb.createCellStyle();
        tableBodyLastRowQuote.cloneStyleFrom(this.tableBodyLastRowStyle);
        if (codeBg != null) {
            tableBodyLastRowQuote.setFillForegroundColor(codeBg);
            tableBodyLastRowQuote.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        this.tableBodyLastRowQuoteStyle = tableBodyLastRowQuote;

        XSSFCellStyle quoteBody = (XSSFCellStyle) wb.createCellStyle();
        quoteBody.cloneStyleFrom(this.normalStyle);
        if (codeBg != null) {
            quoteBody.setFillForegroundColor(codeBg);
            quoteBody.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        quoteBody.setBorderTop(BorderStyle.NONE);
        quoteBody.setBorderRight(BorderStyle.NONE);
        quoteBody.setBorderBottom(BorderStyle.NONE);
        quoteBody.setBorderLeft(BorderStyle.NONE);
        this.blockQuoteBodyStyle = quoteBody;

        XSSFCellStyle quoteLeft = (XSSFCellStyle) wb.createCellStyle();
        quoteLeft.cloneStyleFrom(this.blockQuoteBodyStyle);
        quoteLeft.setBorderLeft(BorderStyle.THICK);
        XSSFColor blue = new XSSFColor(new Color(0, 112, 192), null);
        quoteLeft.setBorderColor(BorderSide.LEFT, blue);
        this.blockQuoteLeftStyle = quoteLeft;

        // 引用ブロック（空行用）
        XSSFCellStyle quoteBlankBody = (XSSFCellStyle) wb.createCellStyle();
        quoteBlankBody.cloneStyleFrom(this.blankRowStyle);
        if (codeBg != null) {
            quoteBlankBody.setFillForegroundColor(codeBg);
            quoteBlankBody.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        quoteBlankBody.setBorderTop(BorderStyle.NONE);
        quoteBlankBody.setBorderRight(BorderStyle.NONE);
        quoteBlankBody.setBorderBottom(BorderStyle.NONE);
        quoteBlankBody.setBorderLeft(BorderStyle.NONE);
        this.blockQuoteBlankBodyStyle = quoteBlankBody;

        XSSFCellStyle quoteBlankLeft = (XSSFCellStyle) wb.createCellStyle();
        quoteBlankLeft.cloneStyleFrom(this.blockQuoteBlankBodyStyle);
        quoteBlankLeft.setBorderLeft(BorderStyle.THICK);
        quoteBlankLeft.setBorderColor(BorderSide.LEFT, blue);
        this.blockQuoteBlankLeftStyle = quoteBlankLeft;

        XSSFCellStyle quoteHrBody = (XSSFCellStyle) wb.createCellStyle();

        quoteHrBody.cloneStyleFrom(this.blockQuoteBlankBodyStyle);
        quoteHrBody.setBorderBottom(BorderStyle.HAIR);

        this.blockQuoteHorizontalRuleBodyStyle = quoteHrBody;
    }

    private void initCodeBlockFrameStyles(Workbook wb) {
        for (int mask = 1; mask < 16; mask++) {
            CellStyle s = wb.createCellStyle();
            s.cloneStyleFrom(this.codeBlockStyle);

            if ((mask & 1) != 0)
                s.setBorderTop(BorderStyle.THIN);
            if ((mask & 2) != 0)
                s.setBorderBottom(BorderStyle.THIN);
            if ((mask & 4) != 0)
                s.setBorderLeft(BorderStyle.THIN);
            if ((mask & 8) != 0)
                s.setBorderRight(BorderStyle.THIN);

            codeBlockFrameStyles[mask] = s;
        }
    }

    // mask bit: 1=TOP, 2=BOTTOM, 4=LEFT, 8=RIGHT
    public CellStyle codeBlockFrameStyle(int mask) {
        if (mask == 0)
            return codeBlockStyle;
        return codeBlockFrameStyles[mask];
    }

    private static CellStyle createBlockQuoteContentStyle(Workbook wb, CellStyle baseStyle, XSSFColor background) {

        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();

        style.cloneStyleFrom(baseStyle);

        if (background != null) {
            style.setFillForegroundColor(background);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        return style;
    }
}