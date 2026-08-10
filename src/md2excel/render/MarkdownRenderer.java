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
        final String raw;
        final String trimmed;
        final int indent;
        final LineKind kind;

        final int headingLevel;
        final String headingText;

        final boolean endsWithHardBreak;
        final String paragraphText;
        final String listMarkerText;
        final String listContentText;

        // kind == BLOCK_QUOTE のときだけ設定。
        // 引用マーカーを除去した内容を通常と同じ parseContent() で解析した結果。
        final LineInfo quotedContent;

        private LineInfo(String raw, String trimmed, int indent, LineKind kind, int headingLevel, String headingText,
                boolean endsWithHardBreak, String paragraphText, String listMarkerText, String listContentText,
                LineInfo quotedContent) {

            this.raw = raw;
            this.trimmed = trimmed;
            this.indent = indent;
            this.kind = kind;

            this.headingLevel = headingLevel;
            this.headingText = headingText;

            this.endsWithHardBreak = endsWithHardBreak;
            this.paragraphText = paragraphText;
            this.listMarkerText = listMarkerText;
            this.listContentText = listContentText;

            this.quotedContent = quotedContent;
        }

        boolean isTableLike() {
            if (kind == LineKind.TABLE_SEPARATOR || kind == LineKind.TABLE_ROW) {
                return true;
            }

            if (kind == LineKind.BLOCK_QUOTE && quotedContent != null) {
                return quotedContent.kind == LineKind.TABLE_SEPARATOR || quotedContent.kind == LineKind.TABLE_ROW;
            }

            return false;
        }

        static LineInfo parse(String rawLine, RenderState st) {
            String trimmed = rawLine.trim();
            int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);

            // コードブロック中だけは最優先。
            if (st.inCodeBlock) {

                // 引用内コードブロックでは、まず引用 marker を1段剥がす。
                if (st.codeBlockInBlockQuote && trimmed.startsWith(">")) {
                    String innerRaw = stripOneQuoteMarker(rawLine);
                    LineInfo inner = parseCodeBlockContent(innerRaw, st);

                    return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null,
                            inner.endsWithHardBreak, null, null, null, inner);
                }

                return parseCodeBlockContent(rawLine, st);
            }

            // 引用は外側のコンテキストとして扱い、
            // 中身は通常行と同じ classifier に通す。
            if (trimmed.startsWith(">")) {
                String innerRaw = stripOneQuoteMarker(rawLine);
                LineInfo inner = parseContent(innerRaw);

                return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null, inner.endsWithHardBreak,
                        null, null, null, inner);
            }

            return parseContent(rawLine);
        }

        private static LineInfo parseCodeBlockContent(String rawLine, RenderState st) {

            String trimmed = rawLine.trim();
            int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);

            if (MdTextUtil.isClosingCodeFenceLine(trimmed, st.codeFenceMarker, st.codeFenceLength)) {

                return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_FENCE, -1, null, false, null, null, null,
                        null);
            }

            return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_LINE, -1, null, false, null, null, null, null);
        }

        /**
         * 引用かどうかに依存しない、実際の block classifier。
         */
        private static LineInfo parseContent(String rawLine) {
            String trimmed = rawLine.trim();
            int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);
            boolean endsWithHardBreak = hasLineEndHardBreak(rawLine);

            // block quote は container なので、
            // 内側も同じ classifier で再帰的に解析する。
            if (trimmed.startsWith(">")) {
                String innerRaw = stripOneQuoteMarker(rawLine);
                LineInfo inner = parseContent(innerRaw);

                return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null, inner.endsWithHardBreak,
                        null, null, null, inner);
            }

            // code fence
            if (MdTextUtil.isOpeningCodeFenceLine(trimmed)) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_FENCE, -1, null, false, null, null, null,
                        null);
            }

            // blank
            if (trimmed.isEmpty()) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.BLANK, -1, null, false, null, null, null, null);
            }

            // horizontal rule
            if (MdTextUtil.isHorizontalRuleLine(trimmed)) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.HORIZONTAL_RULE, -1, null, false, null, null,
                        null, null);
            }

            // table
            if (MarkdownTable.isTableLine(rawLine)) {
                boolean separator = MarkdownTable.isTableSeparatorLine(trimmed);

                return new LineInfo(rawLine, trimmed, indent, separator ? LineKind.TABLE_SEPARATOR : LineKind.TABLE_ROW,
                        -1, null, false, null, null, null, null);
            }

            // heading
            if (trimmed.startsWith("#")) {
                int level = MdTextUtil.countHeadingLevel(trimmed);
                String text = trimmed.substring(level).trim();
                text = MdTextUtil.stripHeadingClosingHashes(text);
                text = stripLineEndHardBreakMarker(text, rawLine);

                return new LineInfo(rawLine, trimmed, indent, LineKind.HEADING, level, text, endsWithHardBreak, null,
                        null, null, null);
            }

            // bullet
            if (trimmed.length() >= 2) {
                char marker = trimmed.charAt(0);

                if ((marker == '*' || marker == '-' || marker == '+') && Character.isWhitespace(trimmed.charAt(1))) {

                    String content = trimmed.substring(2).trim();
                    content = stripLineEndHardBreakMarker(content, rawLine);

                    return new LineInfo(rawLine, trimmed, indent, LineKind.BULLET_ITEM, -1, null, endsWithHardBreak,
                            content, "・ ", content, null);
                }
            }

            // numbered list
            if (MdTextUtil.isNumberedListLine(trimmed)) {
                int markerEnd = findNumberedListMarkerEnd(trimmed);
                String markerText = trimmed.substring(0, markerEnd).trim() + " ";

                String content = trimmed.substring(markerEnd).trim();
                content = stripLineEndHardBreakMarker(content, rawLine);

                return new LineInfo(rawLine, trimmed, indent, LineKind.NUMBER_ITEM, -1, null, endsWithHardBreak,
                        content, markerText, content, null);
            }

            // normal
            String paragraphText = stripLineEndHardBreakMarker(trimmed, rawLine);

            return new LineInfo(rawLine, trimmed, indent, LineKind.NORMAL, -1, null, endsWithHardBreak, paragraphText,
                    null, null, null);
        }

        private static String stripOneQuoteMarker(String rawLine) {
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

            return (i < rawLine.length()) ? rawLine.substring(i) : "";
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

            // Setext heading は直前の paragraph と現在行をセットで判定する。
            // underline 行自体は Excel 行として出力しない。
            int setextHeadingLevel = ParagraphUtil.getSetextHeadingLevel(para, li);

            if (setextHeadingLevel > 0) {
                ParagraphUtil.flushSetextHeading(para, setextHeadingLevel, ctx);
                para = null;
                continue;
            }

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
            ctx.st.codeBlockInBlockQuote = false;
            ctx.st.codeBlockQuoteStartCol = -1;
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

        int quoteStartCol = calcQuoteStartCol(li.indent, ctx.st);

        LineInfo q = li.quotedContent;

        if (q == null) {
            throw new AssertionError("Quoted content is missing");
        }

        // すでに引用内コードブロック中なら、
        // 通常の block quote 前後処理を通さない。
        if (ctx.st.codeBlockInBlockQuote) {

            switch (q.kind) {
            case CODE_LINE:
                handleQuotedCodeLine(q, quoteStartCol, ctx);
                return;

            case CODE_FENCE:
                handleQuotedCodeFence(q, quoteStartCol, ctx);
                return;

            default:
                throw new AssertionError("Unexpected line inside quoted code block: " + q.kind);
            }
        }

        // quoted code の直後に明示的な `>` 空行がある場合、
        // code block 用の自動空行は追加しない。
        boolean explicitBlankAfterQuotedCode = q.kind == LineKind.BLANK
                && ctx.st.lastContentType == RenderState.ContentType.CODE && ctx.st.lastWasBlockQuote;

        if (!explicitBlankAfterQuotedCode) {
            ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.blankRowStyle);
        }

        ctx.st.ensureAutoBlankBeforeBlockQuoteIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        switch (q.kind) {
        case BLANK:
            handleQuotedBlank(ctx, quoteStartCol);
            break;

        case HORIZONTAL_RULE:
            handleQuotedHorizontalRule(ctx, quoteStartCol);
            break;

        case HEADING:
            handleQuotedHeading(q, quoteStartCol, ctx);
            break;

        case CODE_FENCE:
            handleQuotedCodeFence(q, quoteStartCol, ctx);
            break;

        case TABLE_SEPARATOR:
            ctx.st.afterSkipTableSeparatorLine();
            ctx.st.lastWasBlockQuote = true;
            break;

        case TABLE_ROW:
            MarkdownTable.TableRowRenderResult rr = renderTableRow(q.raw, quoteStartCol, ctx);

            for (int r = rr.firstRowNum; r <= rr.lastRowNum; r++) {
                ctx.st.recordBlockQuoteRow(r, quoteStartCol, -1, RenderState.QuoteRowKind.TABLE);
            }

            ctx.st.lastWasBlockQuote = true;
            break;

        default:
            throw new AssertionError("Quote paragraph line should have been handled by ParagraphUtil: " + q.kind);
        }
    }

    private static void handleTableSeparatorLine(LineInfo li, RenderContext ctx) {
        ctx.st.afterSkipTableSeparatorLine();
    }

    private static void handleTableRow(LineInfo li, RenderContext ctx) {
        renderTableRow(li.raw, calcBlockStartCol(li.indent, ctx.st), ctx);
    }

    private static MarkdownTable.TableRowRenderResult renderTableRow(String tableLine, int firstRowStartCol,
            RenderContext ctx) {

        int tableStartCol;
        if (ctx.st.currentTableHeaderRow < 0) {
            tableStartCol = firstRowStartCol;
            ctx.st.currentTableStartCol = tableStartCol;
        } else {
            tableStartCol = ctx.st.currentTableStartCol;
        }

        boolean isHeader = (ctx.st.currentTableHeaderRow < 0);

        MarkdownTable.TableRowRenderResult rr = MarkdownTable.createTableRows(ctx, tableLine, isHeader, tableStartCol);

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
        return rr;
    }

    private static void handleHeading(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankBeforeHeadingIfNeeded(ctx.sheet, ctx.styles.blankRowStyle);

        CellStyle style = resolveHeadingStyle(li.headingLevel, ctx);

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

        // 通常Markdown空行と同じく、
        // 水平線直後の空行は Excel 行を増やさない。
        if (ctx.st.lastRowType == RenderState.RowType.HORIZONTAL_RULE && ctx.st.lastWasBlockQuote) {

            ctx.st.afterConsumeMarkdownBlankWithoutNewRow();
            ctx.st.lastWasBlockQuote = true;
            return;
        }

        writeQuotedBlankRow(ctx, quoteStartCol);
    }

    private static void handleQuotedHorizontalRule(RenderContext ctx, int quoteStartCol) {

        int previousRowNum = ctx.st.rowIndex - 1;

        boolean reusePreviousQuotedBlank = ctx.st.lastRowType == RenderState.RowType.BLANK && previousRowNum >= 0
                && ctx.st.isBlockQuoteRowKind(previousRowNum, RenderState.QuoteRowKind.BLANK);

        Row row;

        if (reusePreviousQuotedBlank) {
            row = ctx.sheet.getRow(previousRowNum);

            if (row == null) {
                row = ctx.sheet.createRow(previousRowNum);
            }

        } else {
            row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.blankRowStyle);

            Cell cell = row.createCell(quoteStartCol);

            MarkdownInline.setResolvedSegmentsCell(ctx.wb, cell, Collections.<MarkdownInline.MdSegment>emptyList(),
                    ctx.styles.blankRowStyle);

            ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, -1, RenderState.QuoteRowKind.HORIZONTAL_RULE);
        }

        ctx.st.afterWriteHorizontalRule();

        // afterWriteHorizontalRule() は通常水平線として
        // lastWasBlockQuote=false にするので引用コンテキストへ戻す。
        ctx.st.lastWasBlockQuote = true;
    }

    private static void handleQuotedHeading(LineInfo q, int quoteStartCol, RenderContext ctx) {

        CellStyle style = resolveHeadingStyle(q.headingLevel, ctx);

        List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(q.headingText);

        if (lines.isEmpty()) {
            lines = Collections.<List<MarkdownInline.MdSegment>>singletonList(
                    Collections.<MarkdownInline.MdSegment>emptyList());
        }

        for (int i = 0; i < lines.size(); i++) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

            Cell cell = row.createCell(quoteStartCol);

            MarkdownInline.setResolvedSegmentsCell(ctx.wb, cell, lines.get(i), style);

            ctx.st.afterWriteQuotedHeading(quoteStartCol);

            ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, quoteStartCol, RenderState.QuoteRowKind.NORMAL);
        }
    }

    private static void handleQuotedCodeFence(LineInfo q, int quoteStartCol, RenderContext ctx) {

        // 開始
        if (!ctx.st.inCodeBlock) {
            ctx.st.currentCodeBlockIndent = q.indent;

            ctx.st.codeFenceMarker = MdTextUtil.getCodeFenceMarker(q.trimmed);

            ctx.st.codeFenceLength = MdTextUtil.getCodeFenceLength(q.trimmed);

            ctx.st.inCodeBlock = true;
            ctx.st.codeBlockInBlockQuote = true;
            ctx.st.codeBlockQuoteStartCol = quoteStartCol;

            ctx.st.lastLineWasTable = false;

            ctx.st.codeBlockFirstRow = -1;
            ctx.st.codeBlockLastRow = -1;
            ctx.st.codeBlockCol = 0;
            ctx.st.codeBlockBaseIndent = -1;

            ctx.st.lastWasBlockQuote = true;
            return;
        }

        // 終了処理・コード枠生成は既存処理を流用する。
        handleCodeFence(q, ctx);

        ctx.st.codeBlockInBlockQuote = false;
        ctx.st.codeBlockQuoteStartCol = -1;

        // handleCodeFence() は通常コードブロックとして終了するため、
        // quote context だけ戻す。
        ctx.st.lastWasBlockQuote = true;
    }

    private static void handleQuotedCodeLine(LineInfo q, int quoteStartCol, RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.CODE_LINE,
                ctx.styles.normalStyle);

        // B列: quote decoration
        // C列: code block frame
        // D列: code text
        int frameStartCol = quoteStartCol;
        int codeCol = clampCol(frameStartCol + 1, ctx.st);

        int leadingSpaces = q.indent;
        int trimSpaces = ctx.st.computeCodeTrimSpaces(leadingSpaces);

        String codeLine = q.raw.substring(trimSpaces);

        Cell cell = row.createCell(codeCol);

        MarkdownInline.setCodeBlockRichTextCell(ctx.wb, cell, codeLine, ctx.styles.codeBlockStyle);

        ctx.st.recordCodeBlockLinePos(row.getRowNum(), frameStartCol);

        // 引用終了時にコードstyleを上書きしないため記録。
        ctx.st.afterWriteCodeLine(codeCol);

        ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, -1, RenderState.QuoteRowKind.CODE);

        ctx.st.lastWasBlockQuote = true;
    }

    private static CellStyle resolveHeadingStyle(int headingLevel, RenderContext ctx) {

        return (headingLevel == 1) ? ctx.styles.heading1Style
                : (headingLevel == 2) ? ctx.styles.heading2Style
                        : (headingLevel == 3) ? ctx.styles.heading3Style : ctx.styles.heading4Style;
    }

    private static void writeQuotedBlankRow(RenderContext ctx, int quoteStartCol) {
        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.blankRowStyle);

        Cell cell = row.createCell(quoteStartCol);
        MarkdownInline.setResolvedSegmentsCell(ctx.wb, cell, Collections.<MarkdownInline.MdSegment>emptyList(),
                ctx.styles.blankRowStyle);

        ctx.st.recordBlockQuoteRow(row.getRowNum(), quoteStartCol, -1, RenderState.QuoteRowKind.BLANK);

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
}