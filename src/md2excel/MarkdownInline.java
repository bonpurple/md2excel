package md2excel;

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

public final class MarkdownInline {

    private MarkdownInline() {
    }

    // Workbook 単位でフォントをキャッシュ（Font は Workbook に紐づくので）
    private static final Map<Workbook, FontCache> FONT_CACHE = Collections.synchronizedMap(new WeakHashMap<>());

    private static final class FontCache {
        final Map<Short, MarkdownFonts> inlineFontsByBaseFontIndex = new HashMap<>();
        final Map<Short, CodeBlockFonts> codeBlockFontsByStyleFontIndex = new HashMap<>();
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

    // インラインMarkdown描画で使うフォントのセット
    private static class MarkdownFonts {
        final Font baseFont;
        final Font boldFont;
        final XSSFFont codeAscii;
        final XSSFFont codeCjk;
        final XSSFFont codeAsciiBold;
        final XSSFFont codeCjkBold;
        boolean baseBold;

        MarkdownFonts(Font baseFont, Font boldFont, XSSFFont codeAscii, XSSFFont codeCjk, XSSFFont codeAsciiBold,
                XSSFFont codeCjkBold) {
            this.baseFont = baseFont;
            this.boldFont = boldFont;
            this.codeAscii = codeAscii;
            this.codeCjk = codeCjk;
            this.codeAsciiBold = codeAsciiBold;
            this.codeCjkBold = codeCjkBold;
        }
    }

    private static class MdSegment {
        final String text;
        final boolean inBold;
        final boolean inCode;

        MdSegment(String text, boolean inBold, boolean inCode) {
            this.text = text;
            this.inBold = inBold;
            this.inCode = inCode;
        }
    }

    // markdownText: **太字** や `code` を含む Markdown 文字列
    // baseStyle : 箇条書き・番号付き・通常テキストなどのセルスタイル
    public static void setMarkdownRichTextCell(Workbook workbook, Cell cell, String markdownText, CellStyle baseStyle) {

        if (markdownText == null) {
            markdownText = "";
        }

        // フォントセット準備
        MarkdownFonts fonts = prepareMarkdownFonts(workbook, baseStyle);
        // Markdown → Segment
        List<MdSegment> segments = parseMarkdownToSegments(markdownText);

        // 空の RichText に追記していく
        XSSFRichTextString rich = new XSSFRichTextString("");
        appendSegmentsToRichText(rich, 0, segments, fonts);

        cell.setCellStyle(baseStyle);
        cell.setCellValue(rich);
    }

    // 既存セルの末尾に Markdown 文字列を追記する。
    // withLeadingSpace == true の場合、追記前に半角スペースを 1 個挿入する。
    public static void appendMarkdownToCell(Workbook workbook, Cell cell, String markdownText, CellStyle baseStyle,
            boolean withLeadingSpace) {

        if (markdownText == null || markdownText.isEmpty()) {
            return;
        }

        // フォントセット
        MarkdownFonts fonts = prepareMarkdownFonts(workbook, baseStyle);

        // Markdown → Segment
        List<MdSegment> segments = parseMarkdownToSegments(markdownText);

        // 共有されているかもしれない RichTextString は直接いじらない
        XSSFRichTextString original = (XSSFRichTextString) cell.getRichStringCellValue();

        // 1) original の内容とフォーマットを完全コピーした新インスタンスを作る
        XSSFRichTextString rich = cloneRichTextString(original);

        // 2) セルには、このクローンだけを持たせる
        cell.setCellValue(rich);

        // 3) 以降の処理はこの rich に対してだけ行う
        String existing = rich.getString();
        int pos = existing.length();

        // 必要なら先頭に半角スペースを追加
        if (withLeadingSpace && pos > 0) {
            rich.append(" ");
            rich.applyFont(pos, pos + 1, fonts.baseFont);
            pos++;
        }

        // Segment を末尾に追記してフォント適用
        appendSegmentsToRichText(rich, pos, segments, fonts);

        cell.setCellStyle(baseStyle);
        cell.setCellValue(rich);
    }

    // Markdown文字列を Segment のリストに分解する共通処理
    private static List<MdSegment> parseMarkdownToSegments(String markdownText) {
        List<MdSegment> segments = new ArrayList<>();
        if (markdownText == null || markdownText.isEmpty()) {
            return segments;
        }

        StringBuilder current = new StringBuilder();
        boolean inBold = false;
        boolean inCode = false;
        int len = markdownText.length();

        for (int i = 0; i < len;) {
            char ch = markdownText.charAt(i);

            // **bold** の開始/終了（コード中では無視）
            if (!inCode && ch == '*' && i + 1 < len && markdownText.charAt(i + 1) == '*') {

                // 本物の太字マーカーかどうか判定
                if (!isRealBoldMarker(markdownText, i, inBold)) {
                    // マーカー扱いしない → 「**」そのものをテキストとして追加
                    current.append("**");
                    i += 2;
                    continue;
                }

                // ここから本物のマーカーとして扱う
                if (current.length() > 0) {
                    segments.add(new MdSegment(current.toString(), inBold, inCode));
                    current.setLength(0);
                }
                inBold = !inBold;
                i += 2;
                continue;
            }

            // `code` の開始/終了（太字の中でも有効にする）
            if (ch == '`') {
                if (current.length() > 0) {
                    segments.add(new MdSegment(current.toString(), inBold, inCode));
                    current.setLength(0);
                }
                inCode = !inCode;
                i++;
                continue;
            }

            // 通常文字
            current.append(ch);
            i++;
        }

        if (current.length() > 0) {
            segments.add(new MdSegment(current.toString(), inBold, inCode));
        }

        return segments;
    }

    // Markdown描画に使うフォントセットを baseStyle から準備する
    private static MarkdownFonts prepareMarkdownFonts(Workbook wb, CellStyle baseStyle) {
        short key = (short) baseStyle.getFontIndex();

        FontCache c = cache(wb);
        MarkdownFonts cached = c.inlineFontsByBaseFontIndex.get(key);
        if (cached != null) {
            return cached;
        }

        // ---- ここからは “1回だけ” 作る ----
        Font base = wb.getFontAt(baseStyle.getFontIndex());
        boolean baseBold = base.getBold();

        Font bold = wb.createFont();
        bold.setFontName(base.getFontName());
        bold.setFontHeightInPoints(base.getFontHeightInPoints());
        bold.setBold(true);

        // インラインコード（赤）
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

        MarkdownFonts mf = new MarkdownFonts(base, bold, codeAscii, codeCjk, codeAsciiBold, codeCjkBold);
        mf.baseBold = baseBold;

        c.inlineFontsByBaseFontIndex.put(key, mf);
        return mf;
    }

    // segments の内容を rich の末尾に追記し、フォントを適用する。
    // startPos は追記開始位置（通常は rich.getString().length()）。
    // 戻り値は追記後の文字列長（次の startPos として使える）。
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
                // 「このコード部分を太字で出したいか？」
                // 1) seg.inBold == true … Markdown で **`code`** のように明示的に太字指定
                // 2) baseBold == true … 見出しセルなど、ベーススタイル自体が太字
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
            } else if (seg.inBold) {
                rich.applyFont(segStart, segEnd, fonts.boldFont);
            } else {
                rich.applyFont(segStart, segEnd, fonts.baseFont);
            }

            pos = segEnd;
        }

        return pos;
    }

    // "**" が「本物の太字マーカー」かどうかを判定する。
    // text.charAt(pos) == '*' かつ text.charAt(pos+1) == '*' 前提。
    private static boolean isRealBoldMarker(String text, int pos, boolean inBold) {
        int len = text.length();

        char prev = (pos > 0) ? text.charAt(pos - 1) : '\0';
        char next = (pos + 2 < len) ? text.charAt(pos + 2) : '\0';

        if (!inBold) {
            // 「開き側」の判定（太字外で見つかった "**"）
            // 1) TE**, CE**, TW** など：
            // 直前が英大文字 かつ 直後が英数ではない（句読点・スペース・行末）なら
            // VS Code 同様「マーカーではなくリテラル」とみなす。
            if (pos > 0 && Character.isUpperCase(prev) && (pos + 2 >= len || !Character.isLetterOrDigit(next))) {
                return false;
            }

            // 2) 行末ギリギリ（"テキスト**" で後ろに何もない）は開きにしても閉じられないのでマーカー扱いしない
            if (pos + 2 >= len) {
                return false;
            }

            // それ以外は素直に「開きマーカー」とみなす
            return true;
        } else {
            // 「閉じ側」の判定（太字中で見つかった "**"）
            // 基本的に全部「本物の閉じマーカー」として扱う。
            // （TE** のようなケースではそもそも inBold が true になっていないので、ここには来ない）
            return true;
        }
    }

    // src のテキストとフォーマット情報を丸ごとコピーした XSSFRichTextString を作る
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

    // =========================
    // ctx版オーバーロード
    // =========================
    public static void setMarkdownRichTextCell(RenderContext ctx, Cell cell, String markdownText, CellStyle baseStyle) {
        setMarkdownRichTextCell(ctx.wb, cell, markdownText, baseStyle);
    }

    static final class BrSplitResult {
        final List<String> lines; // その場で出力できる行（空は基本入れない）
        final boolean endsWithBr; // 末尾が <br> で終わっている（次入力行へ継続）
        final String carryPrefix; // 末尾 <br> 後に継続するとき先頭に付ける接頭辞（太字継続など）

        BrSplitResult(List<String> lines, boolean endsWithBr, String carryPrefix) {
            this.lines = lines;
            this.endsWithBr = endsWithBr;
            this.carryPrefix = carryPrefix;
        }
    }

    static BrSplitResult splitByBrPreserveFormatting(String markdownText) {
        List<String> out = new ArrayList<>();
        if (markdownText == null)
            markdownText = "";

        StringBuilder cur = new StringBuilder();
        boolean inBold = false;
        boolean inCode = false;

        // 次行の先頭に付く「太字継続用プレフィックス」
        String reopenPrefix = "";

        boolean lastWasBr = false;

        for (int i = 0; i < markdownText.length();) {
            char ch = markdownText.charAt(i);

            // ** の扱い（コード中は無視）
            if (!inCode && ch == '*' && i + 1 < markdownText.length() && markdownText.charAt(i + 1) == '*') {
                if (!isRealBoldMarker(markdownText, i, inBold)) {
                    cur.append("**");
                    i += 2;
                    continue;
                }
                cur.append("**");
                inBold = !inBold;
                i += 2;
                lastWasBr = false;
                continue;
            }

            // ` の扱い
            if (ch == '`') {
                cur.append('`');
                inCode = !inCode;
                i++;
                lastWasBr = false;
                continue;
            }

            if (!inCode) {
                int brLen = MdTextUtil.matchBrTagLen(markdownText, i);
                if (brLen > 0) {
                    // 行を確定（太字が開いていたら閉じる）
                    String line = cur.toString();
                    if (inBold)
                        line = line + "**";

                    String trimmed = line.trim();
                    if (!trimmed.isEmpty())
                        out.add(trimmed);

                    // 次行の builder を用意（太字継続なら reopen）
                    cur.setLength(0);
                    reopenPrefix = inBold ? "**" : "";
                    if (!reopenPrefix.isEmpty())
                        cur.append(reopenPrefix);

                    i += brLen;
                    lastWasBr = true;
                    continue;
                }
            }

            // 通常文字
            cur.append(ch);
            i++;
            lastWasBr = false;
        }

        // 末尾処理：最後の行
        String line = cur.toString();
        if (inBold)
            line = line + "**";
        String trimmed = line.trim();
        if (!trimmed.isEmpty())
            out.add(trimmed);

        boolean endsWithBr = lastWasBr;
        String carryPrefix = endsWithBr ? reopenPrefix : "";

        // 末尾が <br> の場合、上で reopenPrefix だけ入った空行が out に入る可能性があるので除去
        if (endsWithBr && !out.isEmpty()) {
            String last = out.get(out.size() - 1);
            if (last.equals("**") || last.equals("`") || last.isEmpty()) {
                out.remove(out.size() - 1);
            }
        }

        return new BrSplitResult(out, endsWithBr, carryPrefix);
    }

    public static boolean hasBrOutsideInlineCode(String markdownText) {
        BrSplitResult sp = splitByBrPreserveFormatting(markdownText);
        return sp.endsWithBr || sp.lines.size() >= 2;
    }

    /**
     * <br>
     * を「半角スペース」にして連結する。 ※インラインコード中の <br>
     * は split されない（そのまま残る） ※太字継続などは split の出力に従う（結果的に見た目は維持される）
     */
    public static String brToSingleSpace(String markdownText) {
        BrSplitResult sp = splitByBrPreserveFormatting(markdownText);
        if (sp.lines.isEmpty())
            return "";
        // lines は trim 済み想定なので " " join で要求どおりになる
        return String.join(" ", sp.lines);
    }
}
