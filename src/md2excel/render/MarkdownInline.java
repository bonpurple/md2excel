package md2excel.render;

import java.awt.Color;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.WeakHashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import md2excel.markdown.MdTextUtil;

public final class MarkdownInline {

    private MarkdownInline() {
    }

    private static final Map<Workbook, FontCache> FONT_CACHE = Collections
            .synchronizedMap(new WeakHashMap<Workbook, FontCache>());

    private static final class FontCache {
        final Map<Short, MarkdownFonts> inlineFontsByBaseFontIndex = new HashMap<Short, MarkdownFonts>();
        final Map<Short, CodeBlockFonts> codeBlockFontsByStyleFontIndex = new HashMap<Short, CodeBlockFonts>();
    }

    private static FontCache cache(Workbook wb) {
        FontCache c = FONT_CACHE.get(wb);
        if (c == null) {
            c = new FontCache();
            FONT_CACHE.put(wb, c);
        }
        return c;
    }

    private static final class CodeBlockFonts {
        final Font ascii;
        final Font cjk;

        CodeBlockFonts(Font ascii, Font cjk) {
            this.ascii = ascii;
            this.cjk = cjk;
        }
    }

    private static final class MarkdownFonts {
        final Font baseFont;
        final Font boldFont;
        final Font italicFont;
        final Font boldItalicFont;
        final XSSFFont codeAscii;
        final XSSFFont codeCjk;
        final XSSFFont codeAsciiBold;
        final XSSFFont codeCjkBold;
        boolean baseBold;

        MarkdownFonts(Font baseFont, Font boldFont, Font italicFont, Font boldItalicFont, XSSFFont codeAscii,
                XSSFFont codeCjk, XSSFFont codeAsciiBold, XSSFFont codeCjkBold) {
            this.baseFont = baseFont;
            this.boldFont = boldFont;
            this.italicFont = italicFont;
            this.boldItalicFont = boldItalicFont;
            this.codeAscii = codeAscii;
            this.codeCjk = codeCjk;
            this.codeAsciiBold = codeAsciiBold;
            this.codeCjkBold = codeCjkBold;
        }
    }

    // package-private: render パッケージ内から直接使う
    static final class MdSegment {
        final String text;
        final boolean inBold;
        final boolean inItalic;
        final boolean inCode;

        MdSegment(String text, boolean inBold, boolean inItalic, boolean inCode) {
            this.text = text;
            this.inBold = inBold;
            this.inItalic = inItalic;
            this.inCode = inCode;
        }
    }

    private enum InlineTokenType {
        TEXT,
        CODE,
        DELIM
    }

    private enum DelimUseKind {
        OPEN_EM,
        CLOSE_EM,
        OPEN_STRONG,
        CLOSE_STRONG
    }

    private static final class DelimUse {
        final int start;
        final int len;
        final DelimUseKind kind;

        DelimUse(int start, int len, DelimUseKind kind) {
            this.start = start;
            this.len = len;
            this.kind = kind;
        }
    }

    private static final class InlineToken {
        final InlineTokenType type;
        final String text; // TEXT/CODEのみ
        final char marker; // DELIMのみ
        final int originalLen; // DELIMのみ
        final boolean canOpen; // DELIMのみ
        final boolean canClose; // DELIMのみ

        int usedOpenChars;
        int usedCloseChars;

        final List<DelimUse> uses = new ArrayList<DelimUse>();

        private InlineToken(InlineTokenType type, String text, char marker, int originalLen, boolean canOpen,
                boolean canClose) {
            this.type = type;
            this.text = text;
            this.marker = marker;
            this.originalLen = originalLen;
            this.canOpen = canOpen;
            this.canClose = canClose;
        }

        static InlineToken text(String text) {
            return new InlineToken(InlineTokenType.TEXT, text, '\0', 0, false, false);
        }

        static InlineToken code(String normalizedText) {
            return new InlineToken(InlineTokenType.CODE, normalizedText, '\0', 0, false, false);
        }

        static InlineToken delim(char marker, int len, boolean canOpen, boolean canClose) {
            return new InlineToken(InlineTokenType.DELIM, null, marker, len, canOpen, canClose);
        }

        boolean isEmphasisDelimiter(char ch) {
            return type == InlineTokenType.DELIM && marker == ch;
        }

        int remainingChars() {
            return originalLen - usedOpenChars - usedCloseChars;
        }

        int availableForOpen() {
            return remainingChars();
        }

        int availableForClose() {
            return remainingChars();
        }

        void consumeAsOpen(int len) {
            int start = originalLen - usedOpenChars - len;
            uses.add(new DelimUse(start, len, (len == 2) ? DelimUseKind.OPEN_STRONG : DelimUseKind.OPEN_EM));
            usedOpenChars += len;
        }

        void consumeAsClose(int len) {
            int start = usedCloseChars;
            uses.add(new DelimUse(start, len, (len == 2) ? DelimUseKind.CLOSE_STRONG : DelimUseKind.CLOSE_EM));
            usedCloseChars += len;
        }

        void sortUses() {
            Collections.sort(uses, new java.util.Comparator<DelimUse>() {
                @Override
                public int compare(DelimUse a, DelimUse b) {
                    return Integer.compare(a.start, b.start);
                }
            });
        }
    }

    private static final class DelimiterRunInfo {
        final boolean canOpen;
        final boolean canClose;

        DelimiterRunInfo(boolean canOpen, boolean canClose) {
            this.canOpen = canOpen;
            this.canClose = canClose;
        }
    }

    // ** / * / _ / `code` のみ（~~ は CommonMark core 非対応なので文字列扱い）
    public static void setMarkdownRichTextCell(Workbook workbook, Cell cell, String markdownText, CellStyle baseStyle) {
        if (markdownText == null) {
            markdownText = "";
        }
        List<MdSegment> segments = parseMarkdownToSegments(markdownText);
        setResolvedSegmentsCell(workbook, cell, segments, baseStyle);
    }

    public static void appendMarkdownToCell(Workbook workbook, Cell cell, String markdownText, CellStyle baseStyle,
            boolean withLeadingSpace) {

        if (markdownText == null || markdownText.isEmpty()) {
            return;
        }

        List<MdSegment> segments = parseMarkdownToSegments(markdownText);
        appendResolvedSegmentsToCell(workbook, cell, segments, baseStyle, withLeadingSpace);
    }

    // package-private: Renderer / Table / CellAppendUtil から使う
    static void setResolvedSegmentsCell(Workbook workbook, Cell cell, List<MdSegment> segments, CellStyle baseStyle) {
        if (segments == null) {
            segments = Collections.<MdSegment>emptyList();
        }

        MarkdownFonts fonts = prepareMarkdownFonts(workbook, baseStyle);

        XSSFRichTextString rich = new XSSFRichTextString("");
        appendSegmentsToRichText(rich, 0, segments, fonts);

        cell.setCellStyle(baseStyle);
        cell.setCellValue(rich);
    }

    static void appendResolvedSegmentsToCell(Workbook workbook, Cell cell, List<MdSegment> segments,
            CellStyle baseStyle, boolean withLeadingSpace) {

        if (segments == null || segments.isEmpty()) {
            return;
        }

        MarkdownFonts fonts = prepareMarkdownFonts(workbook, baseStyle);

        XSSFRichTextString original = (XSSFRichTextString) cell.getRichStringCellValue();
        XSSFRichTextString rich = cloneRichTextString(original);

        cell.setCellValue(rich);

        int pos = rich.getString().length();

        if (withLeadingSpace && pos > 0) {
            rich.append(" ");
            rich.applyFont(pos, pos + 1, fonts.baseFont);
            pos++;
        }

        appendSegmentsToRichText(rich, pos, segments, fonts);

        cell.setCellStyle(baseStyle);
        cell.setCellValue(rich);
    }

    private static List<MdSegment> parseMarkdownToSegments(String markdownText) {
        return parseMarkdown(markdownText).segments;
    }

    private static final class ParseResult {
        final List<MdSegment> segments;
        final String carryPrefix;

        ParseResult(List<MdSegment> segments, String carryPrefix) {
            this.segments = segments;
            this.carryPrefix = carryPrefix;
        }
    }

    private static final class EmphasisState {
        int boldDepth;
        int italicDepth;

        EmphasisState(int boldDepth, int italicDepth) {
            this.boldDepth = boldDepth;
            this.italicDepth = italicDepth;
        }
    }

    private static ParseResult parseMarkdown(String markdownText) {
        List<InlineToken> tokens = tokenizeAndResolveInline(markdownText);
        return buildSegments(tokens);
    }

    private static List<InlineToken> tokenizeAndResolveInline(String markdownText) {
        List<InlineToken> tokens = tokenizeInline(markdownText);
        resolveEmphasis(tokens, '*');
        resolveEmphasis(tokens, '_');
        return tokens;
    }

    private static List<InlineToken> tokenizeInline(String markdownText) {
        List<InlineToken> tokens = new ArrayList<InlineToken>();
        if (markdownText == null || markdownText.isEmpty()) {
            return tokens;
        }

        StringBuilder textBuf = new StringBuilder();

        for (int i = 0; i < markdownText.length();) {
            char ch = markdownText.charAt(i);

            // `code`（複数バッククォート含む）
            if (ch == '`') {
                int tickLen = countBackticks(markdownText, i);
                int close = findClosingBackticks(markdownText, i + tickLen, tickLen);
                if (close >= 0) {
                    flushTextToken(tokens, textBuf);

                    String code = markdownText.substring(i + tickLen, close);
                    code = normalizeCodeSpanContent(code);

                    tokens.add(InlineToken.code(code));

                    i = close + tickLen;
                    continue;
                }

                textBuf.append(markdownText, i, i + tickLen);
                i += tickLen;
                continue;
            }

            // * / _ delimiter run
            if (ch == '*' || ch == '_') {
                int runLen = countRun(markdownText, i, ch);
                DelimiterRunInfo info = analyzeDelimiterRun(markdownText, i, runLen, ch);

                if (info.canOpen || info.canClose) {
                    flushTextToken(tokens, textBuf);
                    tokens.add(InlineToken.delim(ch, runLen, info.canOpen, info.canClose));
                } else {
                    appendRepeated(textBuf, ch, runLen);
                }

                i += runLen;
                continue;
            }

            // ~~ は CommonMark core では非対応なので、そのまま文字列として流す
            textBuf.append(ch);
            i++;
        }

        flushTextToken(tokens, textBuf);
        return tokens;
    }

    private static void flushTextToken(List<InlineToken> tokens, StringBuilder textBuf) {
        if (textBuf.length() == 0) {
            return;
        }
        tokens.add(InlineToken.text(textBuf.toString()));
        textBuf.setLength(0);
    }

    private static int countRun(String text, int pos, char ch) {
        int len = 0;
        while (pos + len < text.length() && text.charAt(pos + len) == ch) {
            len++;
        }
        return len;
    }

    private static DelimiterRunInfo analyzeDelimiterRun(String text, int pos, int runLen, char markerChar) {
        char before = (pos > 0) ? text.charAt(pos - 1) : '\0';
        char after = (pos + runLen < text.length()) ? text.charAt(pos + runLen) : '\0';

        boolean beforeWhitespace = (pos == 0) || isUnicodeWhitespace(before);
        boolean afterWhitespace = (pos + runLen >= text.length()) || isUnicodeWhitespace(after);

        boolean beforePunctuation = (pos > 0) && isPunctuationChar(before);
        boolean afterPunctuation = (pos + runLen < text.length()) && isPunctuationChar(after);

        boolean leftFlanking = !afterWhitespace && (!afterPunctuation || beforeWhitespace || beforePunctuation);

        boolean rightFlanking = !beforeWhitespace && (!beforePunctuation || afterWhitespace || afterPunctuation);

        boolean canOpen;
        boolean canClose;

        if (markerChar == '*') {
            canOpen = leftFlanking;
            canClose = rightFlanking;
        } else {
            // '_'
            canOpen = leftFlanking && (!rightFlanking || beforePunctuation);
            canClose = rightFlanking && (!leftFlanking || afterPunctuation);
        }

        return new DelimiterRunInfo(canOpen, canClose);
    }

    private static boolean isUnicodeWhitespace(char ch) {
        return Character.isWhitespace(ch) || Character.isSpaceChar(ch);
    }

    private static boolean isPunctuationChar(char ch) {
        if (isAsciiPunctuation(ch)) {
            return true;
        }

        int t = Character.getType(ch);
        return t == Character.CONNECTOR_PUNCTUATION || t == Character.DASH_PUNCTUATION
                || t == Character.START_PUNCTUATION || t == Character.END_PUNCTUATION
                || t == Character.INITIAL_QUOTE_PUNCTUATION || t == Character.FINAL_QUOTE_PUNCTUATION
                || t == Character.OTHER_PUNCTUATION;
    }

    private static boolean isAsciiPunctuation(char ch) {
        if (ch > 0x7F) {
            return false;
        }
        return (ch >= '!' && ch <= '/') || (ch >= ':' && ch <= '@') || (ch >= '[' && ch <= '`')
                || (ch >= '{' && ch <= '~');
    }

    private static void resolveEmphasis(List<InlineToken> tokens, char marker) {
        List<Integer> openerStack = new ArrayList<Integer>();

        for (int i = 0; i < tokens.size(); i++) {
            InlineToken closer = tokens.get(i);
            if (!closer.isEmphasisDelimiter(marker)) {
                continue;
            }

            if (closer.canClose) {
                while (closer.availableForClose() > 0) {
                    int openerStackPos = findMatchingEmphasisOpener(tokens, openerStack, closer, marker);
                    if (openerStackPos < 0) {
                        break;
                    }

                    InlineToken opener = tokens.get(openerStack.get(openerStackPos));
                    int useLen = (opener.availableForOpen() >= 2 && closer.availableForClose() >= 2) ? 2 : 1;

                    opener.consumeAsOpen(useLen);
                    closer.consumeAsClose(useLen);

                    if (opener.availableForOpen() == 0) {
                        openerStack.remove(openerStackPos);
                    }
                }
            }

            if (closer.canOpen && closer.availableForOpen() > 0) {
                openerStack.add(i);
            }
        }
    }

    private static int findMatchingEmphasisOpener(List<InlineToken> tokens, List<Integer> openerStack,
            InlineToken closer, char marker) {

        for (int i = openerStack.size() - 1; i >= 0; i--) {
            InlineToken opener = tokens.get(openerStack.get(i));

            if (!opener.isEmphasisDelimiter(marker) || opener.availableForOpen() <= 0) {
                openerStack.remove(i);
                continue;
            }

            if (violatesRuleOfThree(opener, closer)) {
                continue;
            }

            return i;
        }

        return -1;
    }

    /**
     * CommonMark の rule of 3: 片方でも「開けて閉じられる」delimiter run のとき、 opener/closer の元の
     * run 長の合計が 3 の倍数で、 かつ両方とも 3 の倍数ではない場合はマッチ禁止。
     */
    private static boolean violatesRuleOfThree(InlineToken opener, InlineToken closer) {
        boolean oneCanBoth = (opener.canOpen && opener.canClose) || (closer.canOpen && closer.canClose);
        if (!oneCanBoth) {
            return false;
        }

        int sum = opener.originalLen + closer.originalLen;
        if ((sum % 3) != 0) {
            return false;
        }

        boolean bothMultipleOf3 = (opener.originalLen % 3 == 0) && (closer.originalLen % 3 == 0);
        return !bothMultipleOf3;
    }

    private static ParseResult buildSegments(List<InlineToken> tokens) {
        List<MdSegment> out = new ArrayList<MdSegment>();
        StringBuilder carry = new StringBuilder();

        EmphasisState state = new EmphasisState(0, 0);

        for (InlineToken token : tokens) {
            if (token.type == InlineTokenType.TEXT) {
                addMergedSegment(out, token.text, state.boldDepth > 0, state.italicDepth > 0, false);
                continue;
            }

            if (token.type == InlineTokenType.CODE) {
                addMergedSegment(out, token.text, state.boldDepth > 0, state.italicDepth > 0, true);
                continue;
            }

            token.sortUses();

            int pos = 0;
            for (DelimUse use : token.uses) {
                if (use.start > pos) {
                    consumeUnmatchedDelimiterRun(out, token, use.start - pos, state, carry);
                }

                switch (use.kind) {
                case CLOSE_EM:
                    if (state.italicDepth > 0) {
                        state.italicDepth--;
                    }
                    break;
                case CLOSE_STRONG:
                    if (state.boldDepth > 0) {
                        state.boldDepth--;
                    }
                    break;
                case OPEN_EM:
                    state.italicDepth++;
                    break;
                case OPEN_STRONG:
                    state.boldDepth++;
                    break;
                default:
                    break;
                }

                pos = use.start + use.len;
            }

            if (pos < token.originalLen) {
                consumeUnmatchedDelimiterRun(out, token, token.originalLen - pos, state, carry);
            }
        }

        return new ParseResult(out, carry.toString());
    }

    private static void consumeUnmatchedDelimiterRun(List<MdSegment> out, InlineToken token, int len,
            EmphasisState state, StringBuilder carry) {
        if (len <= 0) {
            return;
        }

        int remaining = len;

        if (token.canClose) {
            while (remaining >= 2 && state.boldDepth > 0) {
                state.boldDepth--;
                remaining -= 2;
            }
            while (remaining >= 1 && state.italicDepth > 0) {
                state.italicDepth--;
                remaining--;
            }
        }

        if (token.canOpen) {
            while (remaining >= 2) {
                state.boldDepth++;
                carry.append(token.marker).append(token.marker);
                remaining -= 2;
            }
            while (remaining >= 1) {
                state.italicDepth++;
                carry.append(token.marker);
                remaining--;
            }
        }

        if (remaining > 0) {
            addMergedSegment(out, repeatChar(token.marker, remaining), state.boldDepth > 0, state.italicDepth > 0,
                    false);
        }
    }

    private static void addMergedSegment(List<MdSegment> out, String text, boolean inBold, boolean inItalic,
            boolean inCode) {
        if (text == null || text.isEmpty()) {
            return;
        }

        if (!out.isEmpty()) {
            MdSegment last = out.get(out.size() - 1);
            if (last.inBold == inBold && last.inItalic == inItalic && last.inCode == inCode) {
                out.set(out.size() - 1, new MdSegment(last.text + text, inBold, inItalic, inCode));
                return;
            }
        }

        out.add(new MdSegment(text, inBold, inItalic, inCode));
    }

    private static void appendRepeated(StringBuilder sb, char ch, int count) {
        for (int i = 0; i < count; i++) {
            sb.append(ch);
        }
    }

    private static String repeatChar(char ch, int count) {
        if (count <= 0) {
            return "";
        }
        StringBuilder sb = new StringBuilder(count);
        appendRepeated(sb, ch, count);
        return sb.toString();
    }

    private static MarkdownFonts prepareMarkdownFonts(Workbook wb, CellStyle baseStyle) {
        short key = (short) baseStyle.getFontIndex();

        FontCache c = cache(wb);
        MarkdownFonts cached = c.inlineFontsByBaseFontIndex.get(key);
        if (cached != null) {
            return cached;
        }

        Font base = wb.getFontAt(baseStyle.getFontIndex());
        boolean baseBold = base.getBold();

        Font bold = wb.createFont();
        bold.setFontName(base.getFontName());
        bold.setFontHeightInPoints(base.getFontHeightInPoints());
        bold.setBold(true);

        Font italic = wb.createFont();
        italic.setFontName(base.getFontName());
        italic.setFontHeightInPoints(base.getFontHeightInPoints());
        italic.setItalic(true);

        Font boldItalic = wb.createFont();
        boldItalic.setFontName(base.getFontName());
        boldItalic.setFontHeightInPoints(base.getFontHeightInPoints());
        boldItalic.setBold(true);
        boldItalic.setItalic(true);

        XSSFColor inlineRed = new XSSFColor(new Color(180, 0, 0), null);

        XSSFFont codeAscii = (XSSFFont) wb.createFont();
        codeAscii.setFontName("Consolas");
        codeAscii.setFontHeightInPoints(base.getFontHeightInPoints());
        codeAscii.setColor(inlineRed);

        XSSFFont codeCjk = (XSSFFont) wb.createFont();
        codeCjk.setFontName("Meiryo");
        codeCjk.setFontHeightInPoints(base.getFontHeightInPoints());
        codeCjk.setColor(inlineRed);

        XSSFFont codeAsciiBold = (XSSFFont) wb.createFont();
        codeAsciiBold.setFontName("Consolas");
        codeAsciiBold.setFontHeightInPoints(base.getFontHeightInPoints());
        codeAsciiBold.setBold(true);
        codeAsciiBold.setColor(inlineRed);

        XSSFFont codeCjkBold = (XSSFFont) wb.createFont();
        codeCjkBold.setFontName("Meiryo");
        codeCjkBold.setFontHeightInPoints(base.getFontHeightInPoints());
        codeCjkBold.setBold(true);
        codeCjkBold.setColor(inlineRed);

        MarkdownFonts mf = new MarkdownFonts(base, bold, italic, boldItalic, codeAscii, codeCjk, codeAsciiBold,
                codeCjkBold);
        mf.baseBold = baseBold;

        c.inlineFontsByBaseFontIndex.put(key, mf);
        return mf;
    }

    private static int appendSegmentsToRichText(XSSFRichTextString rich, int startPos, List<MdSegment> segments,
            MarkdownFonts fonts) {

        int pos = startPos;
        if (segments == null || segments.isEmpty()) {
            return pos;
        }

        for (MdSegment seg : segments) {
            String text = seg.text;
            if (text == null || text.isEmpty()) {
                continue;
            }

            int segStart = pos;
            rich.append(text);
            int segEnd = segStart + text.length();

            if (seg.inCode) {
                boolean wantBoldCode = seg.inBold || fonts.baseBold;

                int i = 0;
                while (i < text.length()) {
                    int runStart = i;
                    boolean ascii = MdTextUtil.isAsciiLike(text.charAt(i));
                    i++;
                    while (i < text.length() && MdTextUtil.isAsciiLike(text.charAt(i)) == ascii) {
                        i++;
                    }

                    int runLen = i - runStart;
                    int start = segStart + runStart;
                    int end = start + runLen;

                    if (ascii) {
                        rich.applyFont(start, end, wantBoldCode ? fonts.codeAsciiBold : fonts.codeAscii);
                    } else {
                        rich.applyFont(start, end, wantBoldCode ? fonts.codeCjkBold : fonts.codeCjk);
                    }
                }
            } else if (seg.inBold && seg.inItalic) {
                rich.applyFont(segStart, segEnd, fonts.boldItalicFont);
            } else if (seg.inBold) {
                rich.applyFont(segStart, segEnd, fonts.boldFont);
            } else if (seg.inItalic) {
                rich.applyFont(segStart, segEnd, fonts.italicFont);
            } else {
                rich.applyFont(segStart, segEnd, fonts.baseFont);
            }

            pos = segEnd;
        }

        return pos;
    }

    private static XSSFRichTextString cloneRichTextString(XSSFRichTextString src) {
        if (src == null) {
            return new XSSFRichTextString("");
        }

        String text = src.getString();
        XSSFRichTextString dst = new XSSFRichTextString(text);

        int runs = src.numFormattingRuns();
        for (int i = 0; i < runs; i++) {
            int start = src.getIndexOfFormattingRun(i);
            int end = (i + 1 < runs) ? src.getIndexOfFormattingRun(i + 1) : text.length();

            XSSFFont font = src.getFontOfFormattingRun(i);
            if (font != null) {
                dst.applyFont(start, end, font);
            }
        }

        return dst;
    }

    public static void setCodeBlockRichTextCell(Workbook workbook, Cell cell, String codeText,
            CellStyle codeBlockStyle) {

        FontCache c = cache(workbook);

        short key = (short) codeBlockStyle.getFontIndex();
        CodeBlockFonts fonts = c.codeBlockFontsByStyleFontIndex.get(key);
        if (fonts == null) {
            Font baseFont = workbook.getFontAt(codeBlockStyle.getFontIndex());
            short baseFontHeight = baseFont.getFontHeightInPoints();

            Font codeAsciiFont = workbook.createFont();
            codeAsciiFont.setFontName("Consolas");
            codeAsciiFont.setFontHeightInPoints(baseFontHeight);

            Font codeCjkFont = workbook.createFont();
            codeCjkFont.setFontName("Meiryo");
            codeCjkFont.setFontHeightInPoints(baseFontHeight);

            fonts = new CodeBlockFonts(codeAsciiFont, codeCjkFont);
            c.codeBlockFontsByStyleFontIndex.put(key, fonts);
        }

        XSSFRichTextString rich = new XSSFRichTextString(codeText);

        int i = 0;
        while (i < codeText.length()) {
            int runStart = i;
            boolean ascii = MdTextUtil.isAsciiLike(codeText.charAt(i));
            i++;
            while (i < codeText.length() && MdTextUtil.isAsciiLike(codeText.charAt(i)) == ascii) {
                i++;
            }
            int runLen = i - runStart;
            int start = runStart;
            int end = start + runLen;

            rich.applyFont(start, end, ascii ? fonts.ascii : fonts.cjk);
        }

        cell.setCellStyle(codeBlockStyle);
        cell.setCellValue(rich);
    }

    public static void setMarkdownRichTextCell(RenderContext ctx, Cell cell, String markdownText, CellStyle baseStyle) {
        setMarkdownRichTextCell(ctx.wb, cell, markdownText, baseStyle);
    }

    static final class BrSplitResult {
        final List<List<MdSegment>> lines; // markdown文字列ではなく resolved segment 行
        final boolean endsWithBr;
        final String carryPrefix; // 行継続時に次行先頭へ補う未閉じ強調記号

        BrSplitResult(List<List<MdSegment>> lines, boolean endsWithBr, String carryPrefix) {
            this.lines = lines;
            this.endsWithBr = endsWithBr;
            this.carryPrefix = carryPrefix;
        }
    }

    private static final class BrSplitAccumulator {
        final List<List<MdSegment>> lines = new ArrayList<List<MdSegment>>();
        final List<MdSegment> current = new ArrayList<MdSegment>();
        boolean lastWasBr = false;
    }

    static BrSplitResult splitByBrPreserveFormatting(String markdownText) {
        if (markdownText == null) {
            markdownText = "";
        }
        ParseResult parsed = parseMarkdown(markdownText);
        return splitResolvedSegmentsByBr(parsed.segments, parsed.carryPrefix);
    }

    private static BrSplitResult splitResolvedSegmentsByBr(List<MdSegment> segments, String carryPrefix) {
        BrSplitAccumulator acc = new BrSplitAccumulator();

        for (MdSegment seg : segments) {
            appendSegmentWithBrSplit(seg, acc);
        }

        if (!acc.current.isEmpty()) {
            acc.lines.add(new ArrayList<MdSegment>(acc.current));
            acc.current.clear();
        }

        return new BrSplitResult(acc.lines, acc.lastWasBr, carryPrefix);
    }

    private static void appendSegmentWithBrSplit(MdSegment seg, BrSplitAccumulator acc) {
        if (seg == null || seg.text == null || seg.text.isEmpty()) {
            return;
        }

        // インラインコード中の <br> は分割しない
        if (seg.inCode) {
            addMergedSegment(acc.current, seg.text, seg.inBold, seg.inItalic, true);
            acc.lastWasBr = false;
            return;
        }

        String text = seg.text;
        int start = 0;

        for (int i = 0; i < text.length();) {
            int brLen = MdTextUtil.matchBrTagLen(text, i);
            if (brLen > 0) {
                if (i > start) {
                    addMergedSegment(acc.current, text.substring(start, i), seg.inBold, seg.inItalic, false);
                }
                finishCurrentLineAtBr(acc);
                i += brLen;
                start = i;
                continue;
            }
            i++;
        }

        if (start < text.length()) {
            addMergedSegment(acc.current, text.substring(start), seg.inBold, seg.inItalic, false);
            acc.lastWasBr = false;
        }
    }

    private static void finishCurrentLineAtBr(BrSplitAccumulator acc) {
        if (!acc.current.isEmpty()) {
            acc.lines.add(new ArrayList<MdSegment>(acc.current));
            acc.current.clear();
        }
        acc.lastWasBr = true;
    }

    static List<MdSegment> joinLinesWithSingleSpace(BrSplitResult sp) {
        List<MdSegment> out = new ArrayList<MdSegment>();
        if (sp == null || sp.lines.isEmpty()) {
            return out;
        }

        for (int i = 0; i < sp.lines.size(); i++) {
            if (i > 0) {
                addMergedSegment(out, " ", false, false, false);
            }

            List<MdSegment> line = sp.lines.get(i);
            for (int j = 0; j < line.size(); j++) {
                MdSegment seg = line.get(j);
                addMergedSegment(out, seg.text, seg.inBold, seg.inItalic, seg.inCode);
            }
        }

        return out;
    }

    private static String segmentsToPlainText(List<MdSegment> segments) {
        if (segments == null || segments.isEmpty()) {
            return "";
        }

        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < segments.size(); i++) {
            sb.append(segments.get(i).text);
        }
        return sb.toString();
    }

    public static boolean hasBrOutsideInlineCode(String markdownText) {
        BrSplitResult sp = splitByBrPreserveFormatting(markdownText);
        return sp.endsWithBr || sp.lines.size() >= 2;
    }

    // 互換用。書式は落ちるので、新規コードでは joinLinesWithSingleSpace + setResolvedSegmentsCell
    // を使うこと。
    public static String brToSingleSpace(String markdownText) {
        BrSplitResult sp = splitByBrPreserveFormatting(markdownText);
        return segmentsToPlainText(joinLinesWithSingleSpace(sp));
    }

    private static int countBackticks(String text, int pos) {
        int count = 0;
        while (pos + count < text.length() && text.charAt(pos + count) == '`') {
            count++;
        }
        return count;
    }

    private static int findClosingBackticks(String text, int start, int tickLen) {
        for (int i = start; i < text.length();) {
            if (text.charAt(i) != '`') {
                i++;
                continue;
            }
            int runLen = countBackticks(text, i);
            if (runLen == tickLen) {
                return i;
            }
            i += runLen;
        }
        return -1;
    }

    private static String normalizeCodeSpanContent(String code) {
        if (code == null || code.isEmpty()) {
            return code;
        }
        int start = 0;
        int end = code.length();
        boolean leadingSpace = start < end && Character.isWhitespace(code.charAt(start));
        boolean trailingSpace = end > start && Character.isWhitespace(code.charAt(end - 1));
        if (leadingSpace && trailingSpace) {
            int i = start;
            while (i < end && Character.isWhitespace(code.charAt(i))) {
                i++;
            }
            if (i < end) {
                start++;
                end--;
            }
        }
        return code.substring(start, end);
    }
}
