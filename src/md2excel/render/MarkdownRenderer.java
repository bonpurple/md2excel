package md2excel.render;

import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import md2excel.excel.Md2ExcelSheetUtil;
import md2excel.markdown.ListStackUtil;
import md2excel.markdown.MdTextUtil;

public final class MarkdownRenderer {

    static enum QuoteContentKind {
        BLANK,
        BULLET_ITEM,
        NUMBER_ITEM,
        NORMAL
    }

    private static final class QuotedContent {
        final QuoteContentKind kind;
        final String text; // 通常引用本文、または互換用の表示文字列
        final int indent;

        final String listMarkerText;
        final String listContentText;

        QuotedContent(QuoteContentKind kind, String text, int indent, String listMarkerText, String listContentText) {
            this.kind = kind;
            this.text = text;
            this.indent = indent;
            this.listMarkerText = listMarkerText;
            this.listContentText = listContentText;
        }
    }

    static enum LineKind {
        CODE_FENCE(MdBlockBoundary.Policy.CODE_FENCE),
        CODE_LINE(MdBlockBoundary.Policy.NONE), // inCodeBlock中は境界処理しない（従来通り）
        BLANK(MdBlockBoundary.Policy.MARKDOWN_BLANK),
        HORIZONTAL_RULE(MdBlockBoundary.Policy.HORIZONTAL_RULE),
        BLOCK_QUOTE(MdBlockBoundary.Policy.NONE), // 従来 apply していないなら NONE
        TABLE_SEPARATOR(MdBlockBoundary.Policy.TABLE_LINE),
        TABLE_ROW(MdBlockBoundary.Policy.TABLE_LINE),
        HEADING(MdBlockBoundary.Policy.HEADING),
        BULLET_ITEM(MdBlockBoundary.Policy.BULLET_ITEM),
        NUMBER_ITEM(MdBlockBoundary.Policy.NUMBER_ITEM),
        NORMAL(MdBlockBoundary.Policy.NONE); // 従来 apply していないなら NONE

        final MdBlockBoundary.Policy policy;

        LineKind(MdBlockBoundary.Policy policy) {
            this.policy = policy;
        }
    }

    static final class LineInfo {
        final String raw; // 元行（インデント含む）
        final String trimmed; // raw.trim()
        final int indent; // leading spaces/tabs（raw基準）
        final LineKind kind;

        final int headingLevel;
        final String headingText;
        final String quoteText;

        final QuoteContentKind quoteContentKind;
        final int quoteContentIndent;

        final boolean endsWithHardBreak;
        final String paragraphText; // 通常行本文（末尾 hard break 記法除去済み）
        final String listMarkerText; // "・ " / "12. " など
        final String listContentText; // リスト本文のみ

        final String quoteListMarkerText;
        final String quoteListContentText;

        private LineInfo(String raw, String trimmed, int indent, LineKind kind, int headingLevel, String headingText,
                String quoteText, QuoteContentKind quoteContentKind, int quoteContentIndent, boolean endsWithHardBreak,
                String paragraphText, String listMarkerText, String listContentText, String quoteListMarkerText,
                String quoteListContentText) {

            this.raw = raw;
            this.trimmed = trimmed;
            this.indent = indent;
            this.kind = kind;

            this.headingLevel = headingLevel;
            this.headingText = headingText;
            this.quoteText = quoteText;

            this.quoteContentKind = quoteContentKind;
            this.quoteContentIndent = quoteContentIndent;

            this.endsWithHardBreak = endsWithHardBreak;
            this.paragraphText = paragraphText;
            this.listMarkerText = listMarkerText;
            this.listContentText = listContentText;
            this.quoteListMarkerText = quoteListMarkerText;
            this.quoteListContentText = quoteListContentText;
        }

        boolean isTableLike() {
            return kind == LineKind.TABLE_SEPARATOR || kind == LineKind.TABLE_ROW;
        }

        static LineInfo parse(String rawLine, RenderState st) {
            String trimmed = rawLine.trim();
            int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);
            boolean endsWithHardBreak = hasLineEndHardBreak(rawLine);

            // 1) コードブロック中
            if (st.inCodeBlock) {
                if (MdTextUtil.isClosingCodeFenceLine(trimmed, st.codeFenceMarker, st.codeFenceLength)) {
                    return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_FENCE, -1, null, null, null, 0, false,
                            null, null, null, null, null);
                }
                return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_LINE, -1, null, null, null, 0, false, null,
                        null, null, null, null);
            }

            // 2) 開始フェンス
            if (MdTextUtil.isOpeningCodeFenceLine(trimmed)) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_FENCE, -1, null, null, null, 0, false, null,
                        null, null, null, null);
            }

            // 3) blank
            if (trimmed.isEmpty()) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.BLANK, -1, null, null, null, 0, false, null,
                        null, null, null, null);
            }

            // 4) horizontal rule
            if (MdTextUtil.isHorizontalRuleLine(trimmed)) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.HORIZONTAL_RULE, -1, null, null, null, 0, false,
                        null, null, null, null, null);
            }

            // 5) block quote
            if (trimmed.startsWith(">")) {
                QuotedContent qc = parseQuotedContent(rawLine);
                return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null, qc.text, qc.kind,
                        qc.indent, endsWithHardBreak, null, null, null, qc.listMarkerText, qc.listContentText);
            }

            // 6) table
            if (MarkdownTable.isTableLine(rawLine)) {
                boolean sep = MarkdownTable.isTableSeparatorLine(trimmed);
                return new LineInfo(rawLine, trimmed, indent, sep ? LineKind.TABLE_SEPARATOR : LineKind.TABLE_ROW, -1,
                        null, null, null, 0, false, null, null, null, null, null);
            }

            // 7) heading
            if (trimmed.startsWith("#")) {
                int level = MdTextUtil.countHeadingLevel(trimmed);
                String text = trimmed.substring(level).trim();
                text = MdTextUtil.stripHeadingClosingHashes(text);
                text = stripLineEndHardBreakMarker(text, rawLine);

                return new LineInfo(rawLine, trimmed, indent, LineKind.HEADING, level, text, null, null, 0,
                        endsWithHardBreak, null, null, null, null, null);
            }

            // 8) bullet list
            if (trimmed.length() >= 2) {
                char m = trimmed.charAt(0);
                if ((m == '*' || m == '-' || m == '+') && Character.isWhitespace(trimmed.charAt(1))) {
                    String content = trimmed.substring(2).trim();
                    content = stripLineEndHardBreakMarker(content, rawLine);
                    String markerText = "・ ";

                    return new LineInfo(rawLine, trimmed, indent, LineKind.BULLET_ITEM, -1, null, null, null, 0,
                            endsWithHardBreak, content, markerText, content, null, null);
                }
            }

            // 9) numbered list
            if (MdTextUtil.isNumberedListLine(trimmed)) {
                int markerEnd = findNumberedListMarkerEnd(trimmed);
                String markerText = trimmed.substring(0, markerEnd).trim() + " ";
                String content = trimmed.substring(markerEnd).trim();
                content = stripLineEndHardBreakMarker(content, rawLine);

                return new LineInfo(rawLine, trimmed, indent, LineKind.NUMBER_ITEM, -1, null, null, null, 0,
                        endsWithHardBreak, content, markerText, content, null, null);
            }

            // 10) normal
            String paragraphText = stripLineEndHardBreakMarker(trimmed, rawLine);
            return new LineInfo(rawLine, trimmed, indent, LineKind.NORMAL, -1, null, null, null, 0, endsWithHardBreak,
                    paragraphText, null, null, null, null);
        }

        private static QuotedContent parseQuotedContent(String rawLine) {
            int i = 0;
            while (i < rawLine.length()) {
                char ch = rawLine.charAt(i);
                if (ch == ' ' || ch == '\t') {
                    i++;
                    continue;
                }
                break;
            }

            if (i < rawLine.length() && rawLine.charAt(i) == '>') {
                i++;
            }
            if (i < rawLine.length() && rawLine.charAt(i) == ' ') {
                i++;
            }

            String innerRaw = (i < rawLine.length()) ? rawLine.substring(i) : "";
            String innerTrimmed = innerRaw.trim();
            int innerIndent = MdTextUtil.countLeadingSpacesOrTabs(innerRaw);

            if (innerTrimmed.isEmpty()) {
                return new QuotedContent(QuoteContentKind.BLANK, "", innerIndent, null, null);
            }

            if (innerTrimmed.length() >= 2) {
                char m = innerTrimmed.charAt(0);
                if ((m == '*' || m == '-' || m == '+') && Character.isWhitespace(innerTrimmed.charAt(1))) {
                    String content = innerTrimmed.substring(2).trim();
                    content = stripLineEndHardBreakMarker(content, rawLine);
                    String markerText = "・ ";
                    return new QuotedContent(QuoteContentKind.BULLET_ITEM, markerText + content, innerIndent,
                            markerText, content);
                }
            }

            if (MdTextUtil.isNumberedListLine(innerTrimmed)) {
                int markerEnd = findNumberedListMarkerEnd(innerTrimmed);
                String markerText = innerTrimmed.substring(0, markerEnd).trim() + " ";
                String content = innerTrimmed.substring(markerEnd).trim();
                content = stripLineEndHardBreakMarker(content, rawLine);

                return new QuotedContent(QuoteContentKind.NUMBER_ITEM, markerText + content, innerIndent, markerText,
                        content);
            }

            String text = stripLineEndHardBreakMarker(innerTrimmed, rawLine);
            return new QuotedContent(QuoteContentKind.NORMAL, text, innerIndent, null, null);
        }

        private static boolean hasLineEndHardBreak(String rawLine) {
            return MdTextUtil.hasHardLineBreakByBackslash(rawLine) || MdTextUtil.hasHardLineBreakBySpaces(rawLine);
        }

        private static String stripLineEndHardBreakMarker(String text, String rawLine) {
            if (text == null) {
                return null;
            }
            if (MdTextUtil.hasHardLineBreakByBackslash(rawLine)) {
                return MdTextUtil.removeTrailingBackslash(text);
            }
            return text;
        }

        private static int findNumberedListMarkerEnd(String trimmed) {
            if (trimmed == null || trimmed.isEmpty()) {
                return -1;
            }

            int n = trimmed.length();
            int i = 0;

            while (i < n) {
                char ch = trimmed.charAt(i);
                if (ch < '0' || ch > '9') {
                    break;
                }
                i++;
            }

            if (i == 0 || i >= n) {
                return -1;
            }

            char marker = trimmed.charAt(i);
            if (marker != '.' && marker != ')') {
                return -1;
            }
            i++;

            if (i >= n || !Character.isWhitespace(trimmed.charAt(i))) {
                return -1;
            }

            while (i < n && Character.isWhitespace(trimmed.charAt(i))) {
                i++;
            }

            return i;
        }
    }

    public static void render(Iterator<String> it, RenderContext ctx) {
        RenderState st = ctx.st;
        ParagraphBuffer para = null;

        while (it.hasNext()) {
            String rawLine = it.next();
            LineInfo li = LineInfo.parse(rawLine, st);

            // まず open paragraph が継続できるか判定
            if (para != null && ParagraphUtil.canContinue(para, li)) {
                ParagraphUtil.append(para, li);
                continue;
            }

            // 継続できないなら flush
            if (para != null) {
                ParagraphUtil.flush(para, ctx);
                para = null;
            }

            MdBlockBoundary.closeTableIfLeaving(li.isTableLike(), ctx);
            MdBlockBoundary.apply(li.kind.policy, ctx);

            // paragraph 対象行は start して次へ
            if (ParagraphUtil.isParagraphLine(li)) {
                para = ParagraphUtil.start(li, ctx);
                continue;
            }

            switch (li.kind) {
            case CODE_FENCE:
                handleCodeFence(li, ctx);
                break;
            case CODE_LINE:
                handleInCodeBlock(li, ctx);
                break;
            case BLANK:
                handleBlankLine(li, ctx);
                break;
            case HORIZONTAL_RULE:
                handleHorizontalRule(li, ctx);
                break;
            case BLOCK_QUOTE:
                handleBlockQuote(li, ctx); // quote blank のみ到達
                break;
            case TABLE_SEPARATOR:
                handleTableSeparatorLine(li, ctx);
                break;
            case TABLE_ROW:
                handleTableRow(li, ctx);
                break;
            case HEADING:
                handleHeading(li, ctx);
                break;
            case BULLET_ITEM:
            case NUMBER_ITEM:
            case NORMAL:
                throw new AssertionError("Paragraph line should have been handled earlier: " + li.kind);
            default:
                throw new AssertionError("Unhandled LineKind: " + li.kind);
            }
        }

        if (para != null) {
            ParagraphUtil.flush(para, ctx);
        }

        if (st.lastLineWasTable) {
            MarkdownTable.closeTableIfOpen(ctx.sheet, ctx.styles, st);
        }
        BlockQuoteUtil.closeBlockQuoteIfOpen(ctx.sheet, ctx.styles, st);
    }

    private static void handleCodeFence(LineInfo li, RenderContext ctx) {

        // 開始
        if (!ctx.st.inCodeBlock) {
            ctx.st.ensureAutoBlankIfPrevBlockQuote(ctx.sheet, ctx.styles.blankRowStyle);
            ctx.st.currentCodeBlockIndent = li.indent;

            ctx.st.codeFenceMarker = MdTextUtil.getCodeFenceMarker(li.trimmed);
            ctx.st.codeFenceLength = MdTextUtil.getCodeFenceLength(li.trimmed);

            ctx.st.inCodeBlock = true;
            ctx.st.lastLineWasTable = false;

            ctx.st.codeBlockFirstRow = -1;
            ctx.st.codeBlockLastRow = -1;
            ctx.st.codeBlockCol = 0;
            ctx.st.codeBlockBaseIndent = -1;
            return;
        }

        // 終了
        if (ctx.st.codeBlockFirstRow >= 0 && ctx.st.codeBlockLastRow >= 0) {
            int fillEndCol = Math.max(ctx.st.codeBlockCol, ctx.st.lastColIndex);

            for (int r = ctx.st.codeBlockFirstRow; r <= ctx.st.codeBlockLastRow; r++) {
                Row rowObj = ctx.sheet.getRow(r);
                if (rowObj == null)
                    continue;

                for (int c = ctx.st.codeBlockCol; c <= fillEndCol; c++) {
                    Cell cell = rowObj.getCell(c);
                    if (cell == null)
                        cell = rowObj.createCell(c);

                    boolean isTop = (r == ctx.st.codeBlockFirstRow);
                    boolean isBottom = (r == ctx.st.codeBlockLastRow);
                    boolean isLeft = (c == ctx.st.codeBlockCol);
                    boolean isRight = (c == fillEndCol);

                    int mask = 0;
                    if (isTop)
                        mask |= 1;
                    if (isBottom)
                        mask |= 2;
                    if (isLeft)
                        mask |= 4;
                    if (isRight)
                        mask |= 8;

                    cell.setCellStyle(ctx.styles.codeBlockFrameStyle(mask));
                }
            }
        }

        ctx.st.inCodeBlock = false;
        ctx.st.lastLineWasTable = false;

        ctx.st.codeFenceMarker = '\0';
        ctx.st.codeFenceLength = 0;

        ctx.st.codeBlockFirstRow = -1;
        ctx.st.codeBlockLastRow = -1;
        ctx.st.codeBlockCol = 0;
        ctx.st.codeBlockBaseIndent = -1;
    }

    private static void handleInCodeBlock(LineInfo li, RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.CODE_LINE,
                ctx.styles.normalStyle);

        // 引用ブロックと同じ考え方：
        // 装飾（塗りつぶし・罫線）は block 開始列から、
        // 実際のコード本文は 1 列右に置く
        int frameStartCol = calcBlockStartCol(ctx.st.currentCodeBlockIndent, ctx.st);
        int codeCol = clampCol(frameStartCol + 1, ctx.st);

        int leadingSpaces = li.indent;
        int trimSpaces = ctx.st.computeCodeTrimSpaces(leadingSpaces);
        String codeLine = li.raw.substring(trimSpaces);

        Cell cell = row.createCell(codeCol);
        MarkdownInline.setCodeBlockRichTextCell(ctx.wb, cell, codeLine, ctx.styles.codeBlockStyle);

        // 枠線・塗りつぶしは frameStartCol から張る
        ctx.st.recordCodeBlockLinePos(row.getRowNum(), frameStartCol);

        // 直近の内容列としては本文列を保持
        ctx.st.afterWriteCodeLine(codeCol);
    }

    private static void handleBlankLine(LineInfo li, RenderContext ctx) {
        ctx.st.onMarkdownBlankLine(ctx.sheet, ctx.styles.blankRowStyle);
    }

    private static void handleHorizontalRule(LineInfo li, RenderContext ctx) {
        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx, RowUtil.ReuseKind.HORIZONTAL_RULE,
                ctx.styles.blankRowStyle);
        Md2ExcelSheetUtil.createHorizontalRuleRow(ctx.sheet, row, ctx.styles.horizontalRuleStyle, ctx.st.startColIndex,
                ctx.st.mergeLastCol);
        ctx.st.afterWriteHorizontalRule();
    }

    private static void handleBlockQuote(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        int quoteStartCol = calcQuoteStartCol(li.indent, ctx.st);

        switch (li.quoteContentKind) {
        case BLANK:
            handleQuotedBlank(ctx, quoteStartCol);
            break;

        case NORMAL:
        case BULLET_ITEM:
        case NUMBER_ITEM:
            throw new AssertionError(
                    "Quote paragraph line should have been handled by ParagraphUtil: " + li.quoteContentKind);

        default:
            throw new AssertionError("Unhandled QuoteContentKind: " + li.quoteContentKind);
        }
    }

    private static void handleTableSeparatorLine(LineInfo li, RenderContext ctx) {
        ctx.st.afterSkipTableSeparatorLine();
    }

    private static void handleTableRow(LineInfo li, RenderContext ctx) {

        int tableStartCol;
        if (ctx.st.currentTableHeaderRow < 0) {
            tableStartCol = calcBlockStartCol(li.indent, ctx.st);
            ctx.st.currentTableStartCol = tableStartCol;
        } else {
            tableStartCol = ctx.st.currentTableStartCol;
        }

        boolean isHeader = (ctx.st.currentTableHeaderRow < 0);

        MarkdownTable.TableRowRenderResult rr = MarkdownTable.createTableRows(ctx, li.raw, isHeader, tableStartCol);

        if (isHeader) {
            ctx.st.currentTableHeaderRow = rr.firstRowNum;
            ctx.st.currentTableEndCol = rr.lastCol;
            ctx.st.currentTableBodyStartRow = -1;
            ctx.st.currentTableLastBodyRow = -1;
        } else {
            if (ctx.st.currentTableBodyStartRow < 0) {
                ctx.st.currentTableBodyStartRow = rr.firstRowNum;
            }
            ctx.st.currentTableLastBodyRow = rr.lastRowNum;
            if (rr.lastCol > ctx.st.currentTableEndCol) {
                ctx.st.currentTableEndCol = rr.lastCol;
            }
        }

        ctx.st.afterWriteTableRow(tableStartCol);
    }

    private static void handleHeading(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankBeforeHeadingIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        CellStyle style = (li.headingLevel == 1) ? ctx.styles.heading1Style
                : (li.headingLevel == 2) ? ctx.styles.heading2Style
                        : (li.headingLevel == 3) ? ctx.styles.heading3Style : ctx.styles.heading4Style;

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(li.headingText);
        if (lines.isEmpty()) {
            lines = Collections.<List<MarkdownInline.MdSegment>>singletonList(
                    Collections.<MarkdownInline.MdSegment>emptyList());
        }

        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
        Cell cell = row.createCell(rootCol(ctx.st));
        MarkdownInline.setResolvedSegmentsCell(ctx.wb, cell, lines.get(0), style);
        ctx.st.afterWriteHeading();

        for (int i = 1; i < lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(rootCol(ctx.st));
            MarkdownInline.setResolvedSegmentsCell(ctx.wb, c2, lines.get(i), style);
            ctx.st.afterWriteHeading();
        }
    }

    private static int clampCol(int col, RenderState st) {
        if (col < 0)
            return 0;
        if (col >= st.mergeLastCol)
            return st.mergeLastCol - 1;
        return col;
    }

    private static int rootCol(RenderState st) {
        return clampCol(st.startColIndex, st);
    }

    private static int calcBlockStartCol(int indent, RenderState st) {
        if (indent <= 0) {
            return rootCol(st);
        }

        int col;
        if (!st.listStack.isEmpty()) {
            int depth = ListStackUtil.getDepthForIndent(st.listStack, indent);
            col = st.startColIndex + 1 + depth;
        } else {
            int level = indent / 2;
            if (level < 0)
                level = 0;
            col = st.startColIndex + 1 + level;
        }
        return clampCol(col, st);
    }

    private static int calcQuoteStartCol(int indent, RenderState st) {
        return clampCol(calcBlockStartCol(indent, st) + 1, st);
    }

    private static void handleQuotedBlank(RenderContext ctx, int quoteStartCol) {
        ctx.st.resetOnBlockBoundary();
        ctx.st.clearListContext();
        writeQuotedBlankRow(ctx, quoteStartCol);
    }

    private static void writeQuotedBlankRow(RenderContext ctx, int quoteStartCol) {
        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.blankRowStyle);

        Cell cell = row.createCell(quoteStartCol);
        MarkdownInline.setResolvedSegmentsCell(ctx.wb, cell, Collections.<MarkdownInline.MdSegment>emptyList(),
                ctx.styles.blankRowStyle);

        ctx.st.blankBlockQuoteRows.add(row.getRowNum());

        recordQuotedRow(ctx, row.getRowNum(), quoteStartCol, -1);

        ctx.st.lastRowType = RenderState.RowType.BLANK;
        ctx.st.lastLineWasTable = false;
        ctx.st.lastBlankFromMarkdown = false;
        ctx.st.lastBlankRowIndex = -1;
        ctx.st.lastBlankAfterTable = false;

        ctx.st.lastContentType = RenderState.ContentType.NORMAL;
        ctx.st.lastContentCol = quoteStartCol;
        ctx.st.lastContentWasTable = false;

        ctx.st.lastNormalRowIndex = -1;
        ctx.st.lastNormalIndent = -1;
        ctx.st.bulletDetailActive = false;
        ctx.st.lastWasBlockQuote = true;
    }

    private static void recordQuotedRow(RenderContext ctx, int rowNum, int quoteStartCol, int appendCellCol) {
        int quoteDecorCol = clampCol(quoteStartCol - 1, ctx.st);

        if (!ctx.st.inBlockQuote || ctx.st.blockQuoteFirstRow < 0) {
            ctx.st.inBlockQuote = true;
            ctx.st.blockQuoteFirstRow = rowNum;
            ctx.st.blockQuoteCol = quoteDecorCol;
        }

        if (quoteDecorCol < ctx.st.blockQuoteCol) {
            ctx.st.blockQuoteCol = quoteDecorCol;
        }

        ctx.st.blockQuoteLastRow = rowNum;

        if (appendCellCol >= 0) {
            ctx.st.blockQuoteCellRow = rowNum;
            ctx.st.blockQuoteCellCol = appendCellCol;
        } else {
            ctx.st.blockQuoteCellRow = -1;
            ctx.st.blockQuoteCellCol = -1;
        }

        ctx.st.lastWasBlockQuote = true;
    }
}