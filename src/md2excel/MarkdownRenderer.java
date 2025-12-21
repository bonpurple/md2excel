package md2excel;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

final class MarkdownRenderer {

    private enum LineKind {
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

    private static final class LineInfo {
        final String raw; // 元行（インデント含む）
        final String trimmed; // raw.trim()
        final int indent; // leading spaces/tabs（raw基準）
        final LineKind kind;

        // kind別で使う派生値（必要なものだけ埋める）
        final int headingLevel; // kindがHEADINGのときのみ >=1
        final String headingText; // kindがHEADINGのときのみ
        final String quoteText; // kindがBLOCK_QUOTEのときのみ（">"除去済み）
        final String bulletMarkdownText;// kindがBULLET_ITEMのときのみ（"・ "付与済み）

        private LineInfo(String raw, String trimmed, int indent, LineKind kind, int headingLevel, String headingText,
                String quoteText, String bulletMarkdownText) {
            this.raw = raw;
            this.trimmed = trimmed;
            this.indent = indent;
            this.kind = kind;

            this.headingLevel = headingLevel;
            this.headingText = headingText;
            this.quoteText = quoteText;
            this.bulletMarkdownText = bulletMarkdownText;
        }

        boolean isTableLike() {
            return kind == LineKind.TABLE_SEPARATOR || kind == LineKind.TABLE_ROW;
        }

        static LineInfo parse(String rawLine, RenderState st) {
            String trimmed = rawLine.trim();
            int indent = MdTextUtil.countLeadingSpacesOrTabs(rawLine);

            // 1) code fence は inCodeBlock 中でも最優先（閉じるため）
            if (trimmed.startsWith("```")) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_FENCE, -1, null, null, null);
            }

            // 2) code block 中は「全部 code line」
            if (st.inCodeBlock) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.CODE_LINE, -1, null, null, null);
            }

            // 3) blank
            if (trimmed.isEmpty()) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.BLANK, -1, null, null, null);
            }

            // 4) horizontal rule
            if (trimmed.equals("---")) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.HORIZONTAL_RULE, -1, null, null, null);
            }

            // 5) block quote（テーブル判定より先："> |a|b|" は引用扱い）
            if (trimmed.startsWith(">")) {
                String quoteText = trimmed.substring(1).trim();
                return new LineInfo(rawLine, trimmed, indent, LineKind.BLOCK_QUOTE, -1, null, quoteText, null);
            }

            // 6) table（isTableLine をここで1回だけ）
            if (MarkdownTable.isTableLine(rawLine)) {
                boolean sep = MarkdownTable.isTableSeparatorLine(trimmed);
                return new LineInfo(rawLine, trimmed, indent, sep ? LineKind.TABLE_SEPARATOR : LineKind.TABLE_ROW, -1,
                        null, null, null);
            }

            // 7) heading
            if (trimmed.startsWith("#")) {
                int level = MdTextUtil.countHeadingLevel(trimmed);
                String text = trimmed.substring(level).trim();
                return new LineInfo(rawLine, trimmed, indent, LineKind.HEADING, level, text, null, null);
            }

            // 8) list
            if (trimmed.startsWith("* ")) {
                String content = trimmed.substring(2).trim();
                String bulletMd = "・ " + content;
                return new LineInfo(rawLine, trimmed, indent, LineKind.BULLET_ITEM, -1, null, null, bulletMd);
            }

            if (MdTextUtil.isNumberedListLine(trimmed)) {
                return new LineInfo(rawLine, trimmed, indent, LineKind.NUMBER_ITEM, -1, null, null, null);
            }

            return new LineInfo(rawLine, trimmed, indent, LineKind.NORMAL, -1, null, null, null);
        }
    }

    static void render(Iterator<String> it, RenderContext ctx) {
        RenderState st = ctx.st;

        while (it.hasNext()) {
            String rawLine = it.next();
            LineInfo li = LineInfo.parse(rawLine, st);

            MdBlockBoundary.closeTableIfLeaving(li.isTableLike(), ctx);

            // ここで必ず境界処理を実施（呼び忘れが起きない）
            MdBlockBoundary.apply(li.kind.policy, ctx);

            if (tryConsumeHeadingBr(li, ctx))
                continue;
            if (tryConsumeListBr(li, ctx))
                continue;
            if (tryConsumeQuoteBr(li, ctx))
                continue;
            if (tryConsumeSameColBr(li, ctx))
                continue;

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
                handleBlockQuote(li, ctx);
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
                handleBullet(li, ctx);
                break;
            case NUMBER_ITEM:
                handleNumberedList(li, ctx);
                break;
            case NORMAL:
                handleNormalText(li, ctx);
                break;
            default:
                // 新しい LineKind を増やしたのに switch を直してない事故を早期に検出
                throw new AssertionError("Unhandled LineKind: " + li.kind);
            }
        }

        if (st.lastLineWasTable) {
            MarkdownTable.closeTableIfOpen(ctx.sheet, ctx.styles, st);
        }
        BlockQuoteUtil.closeBlockQuoteIfOpen(ctx.sheet, ctx.styles, st);
    }

    // -------------------- handler: ``` --------------------
    private static void handleCodeFence(LineInfo li, RenderContext ctx) {

        if (!ctx.st.inCodeBlock) {
            ctx.st.ensureAutoBlankIfPrevBlockQuote(ctx.sheet, ctx.styles.normalStyle);
            ctx.st.currentCodeBlockIndent = li.indent;
        }

        if (ctx.st.inCodeBlock && ctx.st.codeBlockFirstRow >= 0 && ctx.st.codeBlockLastRow >= 0) {
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

        ctx.st.inCodeBlock = !ctx.st.inCodeBlock;
        ctx.st.lastLineWasTable = false;

        ctx.st.codeBlockFirstRow = -1;
        ctx.st.codeBlockLastRow = -1;
        ctx.st.codeBlockCol = 0;
        ctx.st.codeBlockBaseIndent = -1;
    }

    // -------------------- handler: コードブロック中 --------------------
    private static void handleInCodeBlock(LineInfo li, RenderContext ctx) {

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.CODE_LINE,
                ctx.styles.normalStyle);

        int depth = ListStackUtil.getDepthForIndent(ctx.st.listStack, ctx.st.currentCodeBlockIndent);
        int codeCol = clampCol(1 + depth, ctx.st);

        int leadingSpaces = li.indent;
        int trimSpaces = ctx.st.computeCodeTrimSpaces(leadingSpaces);
        String codeLine = li.raw.substring(trimSpaces);

        Cell cell = row.createCell(codeCol);
        MarkdownInline.setCodeBlockRichTextCell(ctx.wb, cell, codeLine, ctx.styles.codeBlockStyle);

        ctx.st.recordCodeBlockLinePos(row.getRowNum(), codeCol);
        ctx.st.afterWriteCodeLine(codeCol);
    }

    // -------------------- handler: 空行 --------------------
    private static void handleBlankLine(LineInfo li, RenderContext ctx) {
        ctx.st.onMarkdownBlankLine(ctx.sheet, ctx.styles.normalStyle);
    }

    // -------------------- handler: 水平線（---） --------------------
    private static void handleHorizontalRule(LineInfo li, RenderContext ctx) {
        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx, RowUtil.ReuseKind.HORIZONTAL_RULE,
                ctx.styles.normalStyle);
        Md2ExcelSheetUtil.createHorizontalRuleRow(ctx.sheet, row, ctx.styles.horizontalRuleStyle, ctx.st.mergeLastCol);
        ctx.st.afterWriteHorizontalRule();
    }

    // -------------------- handler: 引用（>） --------------------
    private static void handleBlockQuote(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankIfPrevCodeBlock(ctx.sheet, ctx.styles.normalStyle);

        // 必ず split API を通す（ここが差分ポイント）
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(li.quoteText);
        boolean hasBr = sp.endsWithBr || sp.lines.size() >= 2;

        // 既存の「連続引用→同一セル追記」は <br> がないときだけ
        if (!hasBr && ctx.st.inBlockQuote && ctx.st.blockQuoteCellRow >= 0 && ctx.st.blockQuoteCellCol >= 0
                && ctx.st.lastRowType == RenderState.RowType.OTHER && !ctx.st.lastBlankFromMarkdown) {

            // split した結果（1行）を使っても良いが、差分最小でそのまま
            CellAppendUtil.appendMarkdownWithSpace(ctx, ctx.st.blockQuoteCellRow, ctx.st.blockQuoteCellCol,
                    li.quoteText, ctx.styles.normalStyle);

            ctx.st.afterAppendBlockQuoteLine();
            return;
        }

        int depth = ListStackUtil.getDepthForIndent(ctx.st.listStack, li.indent);
        int col = clampCol(1 + depth, ctx.st);

        // <br> 分割して同じ列に縦展開（sp をそのまま使う）
        // 1行目
        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.BLOCK_QUOTE,
                ctx.styles.normalStyle);
        Cell cell = row.createCell(col);
        MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, sp.lines.isEmpty() ? "" : sp.lines.get(0),
                ctx.styles.normalStyle);
        ctx.st.afterWriteBlockQuoteLine(row.getRowNum(), col);

        // 2行目以降（同列）
        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(col);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, c2, sp.lines.get(i), ctx.styles.normalStyle);
            ctx.st.afterWriteBlockQuoteLine(r2.getRowNum(), col);
        }

        // 行末 <br> は次入力行へ継続（次行が > でも通常行でも吸う）
        ctx.st.pendingQuoteBr = sp.endsWithBr;
        ctx.st.pendingQuoteBrCol = col;
        ctx.st.pendingQuoteBrCarry = sp.carryPrefix;
    }

    private static boolean tryConsumeQuoteBr(LineInfo li, RenderContext ctx) {
        if (!ctx.st.pendingQuoteBr)
            return false;

        // BLOCK_QUOTE または NORMAL を継続として扱う
        if (li.kind != LineKind.BLOCK_QUOTE && li.kind != LineKind.NORMAL) {
            ctx.st.pendingQuoteBr = false;
            ctx.st.pendingQuoteBrCarry = "";
            return false;
        }

        String text = (li.kind == LineKind.BLOCK_QUOTE) ? li.quoteText : li.trimmed;
        text = ctx.st.pendingQuoteBrCarry + text;

        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        for (String line : sp.lines) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(ctx.st.pendingQuoteBrCol);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, line, ctx.styles.normalStyle);
            ctx.st.afterWriteBlockQuoteLine(row.getRowNum(), ctx.st.pendingQuoteBrCol);
        }

        ctx.st.pendingQuoteBr = sp.endsWithBr;
        ctx.st.pendingQuoteBrCarry = sp.carryPrefix;
        return true;
    }

    // -------------------- handler: テーブル区切り（|---|---|） --------------------
    private static void handleTableSeparatorLine(LineInfo li, RenderContext ctx) {

        ctx.st.afterSkipTableSeparatorLine();
    }

    // -------------------- handler: テーブル行（| a | b |） --------------------
    private static void handleTableRow(LineInfo li, RenderContext ctx) {

        int tableStartCol;
        if (ctx.st.currentTableHeaderRow < 0) {
            int depth = ListStackUtil.getDepthForIndent(ctx.st.listStack, li.indent);
            tableStartCol = clampCol(1 + depth, ctx.st);
            ctx.st.currentTableStartCol = tableStartCol;
        } else {
            tableStartCol = ctx.st.currentTableStartCol;
        }

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.TABLE_ROW,
                ctx.styles.normalStyle);

        boolean isHeader = (ctx.st.currentTableHeaderRow < 0);
        int lastCol = MarkdownTable.createTableRow(ctx.wb, li.raw, row, ctx.styles, isHeader, tableStartCol);

        int rowNum = row.getRowNum();
        if (isHeader) {
            ctx.st.currentTableHeaderRow = rowNum;
            ctx.st.currentTableEndCol = lastCol;
            ctx.st.currentTableBodyStartRow = -1;
            ctx.st.currentTableLastBodyRow = -1;
        } else {
            if (ctx.st.currentTableBodyStartRow < 0)
                ctx.st.currentTableBodyStartRow = rowNum;
            ctx.st.currentTableLastBodyRow = rowNum;
            if (lastCol > ctx.st.currentTableEndCol)
                ctx.st.currentTableEndCol = lastCol;
        }

        ctx.st.afterWriteTableRow(tableStartCol);
    }

    // -------------------- handler: 見出し --------------------
    private static void handleHeading(LineInfo li, RenderContext ctx) {
        ctx.st.ensureAutoBlankBeforeHeadingIfNeeded(ctx.sheet, ctx.styles.normalStyle);

        CellStyle style = (li.headingLevel == 1) ? ctx.styles.heading1Style
                : (li.headingLevel == 2) ? ctx.styles.heading2Style
                        : (li.headingLevel == 3) ? ctx.styles.heading3Style : ctx.styles.heading4Style;

        // 必ず split API を通す
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(li.headingText);

        // 1行目
        Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
        Cell cell = row.createCell(0);
        MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, sp.lines.isEmpty() ? "" : sp.lines.get(0), style);
        ctx.st.afterWriteHeading();

        // 2行目以降（同じ見出しスタイルで A列に縦展開）
        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(0);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, c2, sp.lines.get(i), style);
            ctx.st.afterWriteHeading();
        }

        // 行末が <br> なら次入力行へ継続（太字継続も carry で保持）
        ctx.st.pendingHeadingBr = sp.endsWithBr;
        ctx.st.pendingHeadingLevel = li.headingLevel;
        ctx.st.pendingHeadingCarry = sp.carryPrefix;
    }

    private static boolean tryConsumeHeadingBr(LineInfo li, RenderContext ctx) {
        if (!ctx.st.pendingHeadingBr)
            return false;
        if (li.kind != LineKind.NORMAL)
            return false; // 例の通り、次行が普通行のときだけ継続

        CellStyle style = (ctx.st.pendingHeadingLevel == 1) ? ctx.styles.heading1Style
                : (ctx.st.pendingHeadingLevel == 2) ? ctx.styles.heading2Style
                        : (ctx.st.pendingHeadingLevel == 3) ? ctx.styles.heading3Style : ctx.styles.heading4Style;

        String text = ctx.st.pendingHeadingCarry + li.trimmed;
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        for (String line : sp.lines) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(0);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, line, style);
            ctx.st.afterWriteHeading();
        }

        ctx.st.pendingHeadingBr = sp.endsWithBr;
        ctx.st.pendingHeadingCarry = sp.carryPrefix;
        return true;
    }

    // -------------------- handler: 箇条書き --------------------
    private static void handleBullet(LineInfo li, RenderContext ctx) {
        int depth = ListStackUtil.updateListDepth(ctx.st.listStack, li.indent, false);
        int col = clampCol(1 + depth, ctx.st);

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.BULLET_ITEM,
                ctx.styles.normalStyle);

        // 必ず split API を通す（太字は維持）
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(li.bulletMarkdownText);

        // 1行目：B列
        Cell cell = row.createCell(col);
        String first = sp.lines.isEmpty() ? "" : sp.lines.get(0);
        MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, first, ctx.styles.bulletStyle);
        ctx.st.afterWriteBulletItem(row.getRowNum(), col);

        // 2行目以降：次行 & 1つ右（C列以降）
        int contCol = clampCol(col + 1, ctx.st);
        int lastRowNum = row.getRowNum();
        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(contCol);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, c2, sp.lines.get(i), ctx.styles.bulletStyle);
            lastRowNum = r2.getRowNum();
            ctx.st.afterWriteNormalText(lastRowNum, contCol, 0, false);
        }

        // 継続状態をセット（行末 <br> や、2行目以降があるなら「以後 NORMAL を右セルに追記」）
        boolean needCont = sp.endsWithBr || sp.lines.size() >= 2;
        if (needCont) {
            ctx.st.pendingListBr = true;
            ctx.st.pendingListBrCol = contCol;
            ctx.st.pendingListBrRow = (sp.lines.size() >= 2) ? lastRowNum : -1; // まだセル未作成なら -1
            ctx.st.pendingListBrHasCell = (sp.lines.size() >= 2);
            ctx.st.pendingListBrStyle = ctx.styles.bulletStyle;
            ctx.st.pendingListBrCarry = sp.carryPrefix;

            // 既存の「説明行追記」ロジックを止める（ここ重要）
            ctx.st.bulletDetailActive = false;
        }
    }

    // -------------------- handler: 番号付き --------------------
    private static void handleNumberedList(LineInfo li, RenderContext ctx) {
        int depth = ListStackUtil.updateListDepth(ctx.st.listStack, li.indent, true);
        int col = clampCol(1 + depth, ctx.st);

        Row row = RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st, RowUtil.ReuseKind.NUMBER_ITEM,
                ctx.styles.normalStyle);

        // 必ず split API を通す
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(li.trimmed);

        Cell cell = row.createCell(col);
        String first = sp.lines.isEmpty() ? "" : sp.lines.get(0);
        MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, first, ctx.styles.listStyle);
        ctx.st.afterWriteNumberedItem(li.indent, col);

        int contCol = clampCol(col + 1, ctx.st);
        int lastRowNum = row.getRowNum();
        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(contCol);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, c2, sp.lines.get(i), ctx.styles.listStyle);
            lastRowNum = r2.getRowNum();
            ctx.st.afterWriteNormalText(lastRowNum, contCol, 0, false);
        }

        boolean needCont = sp.endsWithBr || sp.lines.size() >= 2;
        if (needCont) {
            ctx.st.pendingListBr = true;
            ctx.st.pendingListBrCol = contCol;
            ctx.st.pendingListBrRow = (sp.lines.size() >= 2) ? lastRowNum : -1;
            ctx.st.pendingListBrHasCell = (sp.lines.size() >= 2);
            ctx.st.pendingListBrStyle = ctx.styles.listStyle;
            ctx.st.pendingListBrCarry = sp.carryPrefix;

            // 既存の番号付き「説明行」追記を止める
            ctx.st.inNestedNumberBlock = false;
        }
    }

    private static boolean tryConsumeListBr(LineInfo li, RenderContext ctx) {
        if (!ctx.st.pendingListBr)
            return false;

        // NORMAL 以外が来たら継続終了（例：次に見出し/別リスト/空行など）
        if (li.kind != LineKind.NORMAL) {
            ctx.st.pendingListBr = false;
            ctx.st.pendingListBrHasCell = false;
            ctx.st.pendingListBrRow = -1;
            ctx.st.pendingListBrCarry = "";
            return false;
        }

        String text = ctx.st.pendingListBrCarry + li.trimmed;
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        // まだセルが無い（* test1<br> の直後など）なら、ここで「次行」を作って右セルに書く
        if (!ctx.st.pendingListBrHasCell) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(ctx.st.pendingListBrCol);
            String first = sp.lines.isEmpty() ? "" : sp.lines.get(0);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, first, ctx.st.pendingListBrStyle);

            ctx.st.pendingListBrRow = row.getRowNum();
            ctx.st.pendingListBrHasCell = true;

            // 2行目以降（<br> がこの行にもあった場合）は同列に縦展開
            for (int i = 1; i < sp.lines.size(); i++) {
                Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
                Cell c2 = r2.createCell(ctx.st.pendingListBrCol);
                MarkdownInline.setMarkdownRichTextCell(ctx.wb, c2, sp.lines.get(i), ctx.st.pendingListBrStyle);
                ctx.st.pendingListBrRow = r2.getRowNum();
            }
        } else {
            // 既存セルに追記（半角スペース付き）
            if (!sp.lines.isEmpty()) {
                CellAppendUtil.appendMarkdownWithSpace(ctx, ctx.st.pendingListBrRow, ctx.st.pendingListBrCol,
                        sp.lines.get(0), ctx.st.pendingListBrStyle);
            }
            // 2行目以降は新規行（同列）
            for (int i = 1; i < sp.lines.size(); i++) {
                Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
                Cell c2 = r2.createCell(ctx.st.pendingListBrCol);
                MarkdownInline.setMarkdownRichTextCell(ctx.wb, c2, sp.lines.get(i), ctx.st.pendingListBrStyle);
                ctx.st.pendingListBrRow = r2.getRowNum();
            }
        }

        ctx.st.pendingListBrCarry = sp.carryPrefix;
        // 継続は「空行か他ブロックが来るまで」続ける（例3に合う）
        return true;
    }

    // -------------------- handler: 通常テキスト --------------------
    private static void handleNormalText(LineInfo li, RenderContext ctx) {

        // ---- 0) 直前の「行末 <br>」継続を最優先で消費（同じ列に次行出力） ----
        if (ctx.st.pendingSameColBr) {
            consumePendingSameColBr(li, ctx);
            return;
        }

        // ---- 1) 引用セルへの追記（ただし <br> が無い場合だけ） ----
        // <br> があると「次行セルへ展開」が必要になるため、ここでは追記しない
        if (!MarkdownInline.hasBrOutsideInlineCode(li.trimmed) && tryAppendToOpenBlockQuote(li.trimmed, ctx)) {
            return;
        }

        ctx.st.ensureAutoBlankIfPrevHeading(ctx.sheet, ctx.styles.normalStyle);

        int indent = li.indent;

        // ---- 2) 既存セルへの追記（箇条書き説明 / 番号付き説明 / 同インデント連結） ----
        // ここで true を返した場合、tryAppendToPrevCells 側で <br> 対応している想定
        if (tryAppendToPrevCells(ctx, li.trimmed, indent)) {
            return;
        }

        // ---- 3) 新規セルに通常出力（<br> は縦展開） ----
        NormalTextFlags f = buildNormalTextFlags(indent, ctx.st);
        int col = calcNormalTextCol(indent, ctx.st, f);

        boolean reuseBlank = shouldReuseBlankForNormalText(indent, ctx.st, f);
        Row row = reuseBlank ? RowUtil.reuseLastMarkdownBlankRow(ctx.sheet, ctx.st, ctx.styles.normalStyle)
                : RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

        // <br> を分割（インラインコード中は分割しない / 太字は維持）
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(li.trimmed);

        // 1行目
        Cell cell = row.createCell(col);
        String first = sp.lines.isEmpty() ? "" : sp.lines.get(0);
        MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, first, ctx.styles.normalStyle);
        ctx.st.afterWriteNormalText(row.getRowNum(), col, indent, f.isListNote);

        // 2行目以降：同じ列に新規行で縦展開
        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(col);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, c2, sp.lines.get(i), ctx.styles.normalStyle);

            // ここは「追加行」なので isListNote は false（inListBlock を二重に落とさない）
            ctx.st.afterWriteNormalText(r2.getRowNum(), col, indent, false);
        }

        // 行末が <br> なら次入力行を「同じ列の次行」に継続
        if (sp.endsWithBr) {
            ctx.st.pendingSameColBr = true;
            ctx.st.pendingSameColBrCol = col;
            ctx.st.pendingSameColBrStyle = ctx.styles.normalStyle;
            ctx.st.pendingSameColBrCarry = sp.carryPrefix;
        }
    }

    /**
     * 直前行が「...<br>
     * 」で終わっていた場合、次の入力行（NORMAL）を 同じ列の次行に出力する（さらに <br>
     * があれば縦展開し、行末 <br>
     * なら継続する）
     */
    private static void consumePendingSameColBr(LineInfo li, RenderContext ctx) {
        int col = ctx.st.pendingSameColBrCol;
        CellStyle style = ctx.st.pendingSameColBrStyle;

        String text = ctx.st.pendingSameColBrCarry + li.trimmed;
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        // 継続なので必ず「新しい行」に出す
        for (String line : sp.lines) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(col);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, line, style);

            // indent は 0 扱いでOK（列は固定で出しているため）
            ctx.st.afterWriteNormalText(row.getRowNum(), col, 0, false);
        }

        // 継続更新
        ctx.st.pendingSameColBr = sp.endsWithBr;
        ctx.st.pendingSameColBrCarry = sp.carryPrefix;

        if (!ctx.st.pendingSameColBr) {
            ctx.st.pendingSameColBrCol = 0;
            ctx.st.pendingSameColBrStyle = null;
            ctx.st.pendingSameColBrCarry = "";
        }
    }

    private static boolean tryConsumeSameColBr(LineInfo li, RenderContext ctx) {
        if (!ctx.st.pendingSameColBr)
            return false;
        if (li.kind != LineKind.NORMAL) {
            ctx.st.pendingSameColBr = false;
            ctx.st.pendingSameColBrCarry = "";
            return false;
        }

        String text = ctx.st.pendingSameColBrCarry + li.trimmed;
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(text);

        for (String line : sp.lines) {
            Row row = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell cell = row.createCell(ctx.st.pendingSameColBrCol);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, cell, line, ctx.st.pendingSameColBrStyle);
            ctx.st.afterWriteNormalText(row.getRowNum(), ctx.st.pendingSameColBrCol, 0, false);
        }

        ctx.st.pendingSameColBr = sp.endsWithBr;
        ctx.st.pendingSameColBrCarry = sp.carryPrefix;
        return true;
    }

    private static boolean tryAppendToOpenBlockQuote(String trimmed, RenderContext ctx) {
        if (!ctx.st.inBlockQuote || ctx.st.blockQuoteCellRow < 0 || ctx.st.blockQuoteCellCol < 0)
            return false;
        if (ctx.st.lastRowType != RenderState.RowType.OTHER || ctx.st.lastBlankFromMarkdown)
            return false;

        CellAppendUtil.appendMarkdownWithSpace(ctx.sheet, ctx.wb, ctx.styles, ctx.st.blockQuoteCellRow,
                ctx.st.blockQuoteCellCol, trimmed, ctx.styles.normalStyle);

        ctx.st.afterAppendToOpenBlockQuoteFromNormalText(ctx.st.blockQuoteCellRow, ctx.st.blockQuoteCellCol);
        return true;
    }

    private static boolean tryAppendToPrevCells(RenderContext ctx, String trimmed, int indent) {

        // 1) 箇条書き説明行（直前 bullet セルに追記）
        if (ctx.st.bulletDetailActive && indent > 0 && ctx.st.bulletDetailRow == ctx.st.rowIndex - 1) {
            appendToExistingCellWithBr(ctx, ctx.st.bulletDetailRow, ctx.st.bulletDetailCol, trimmed,
                    ctx.styles.bulletStyle, indent);
            return true;
        }

        // 2) 番号付き説明行（直前 numbered セルに追記）
        boolean isNumberDetail = ctx.st.rowIndex > 0 && ctx.st.inNestedNumberBlock
                && ctx.st.lastContentType == RenderState.ContentType.NUMBER && indent > ctx.st.nestedNumberIndent
                && ctx.st.lastRowType == RenderState.RowType.OTHER && !ctx.st.lastBlankFromMarkdown;

        if (isNumberDetail) {
            int rowNum = ctx.st.rowIndex - 1;
            appendToExistingCellWithBr(ctx, rowNum, ctx.st.nestedNumberCol, trimmed, ctx.styles.listStyle, indent);
            return true;
        }

        // 3) 同インデント連結（直前 normal セルに追記）
        boolean isSameIndentConcat = ctx.st.lastContentType == RenderState.ContentType.NORMAL
                && ctx.st.lastNormalRowIndex >= 0 && ctx.st.lastNormalIndent == indent
                && ctx.st.lastRowType == RenderState.RowType.OTHER && !ctx.st.lastBlankFromMarkdown
                && ctx.st.lastNormalRowIndex == ctx.st.rowIndex - 1;

        if (isSameIndentConcat) {
            int rowNum = ctx.st.lastNormalRowIndex;
            int colNum = ctx.st.lastContentCol;
            appendToExistingCellWithBr(ctx, rowNum, colNum, trimmed, ctx.styles.normalStyle, indent);
            return true;
        }

        return false;
    }

    private static void appendToExistingCellWithBr(RenderContext ctx, int targetRow, int targetCol, String markdown,
            CellStyle style, int indent) {

        // 必ず split API を通す
        MarkdownInline.BrSplitResult sp = MarkdownInline.splitByBrPreserveFormatting(markdown);

        // 1行目：既存セルへ追記（従来通り）
        if (!sp.lines.isEmpty()) {
            CellAppendUtil.appendMarkdownWithSpace(ctx, targetRow, targetCol, sp.lines.get(0), style);
            ctx.st.afterAppendNormalToExistingCell(targetRow, targetCol, indent);
        }

        // 2行目以降：次行（同列）に新規出力
        for (int i = 1; i < sp.lines.size(); i++) {
            Row r2 = RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);
            Cell c2 = r2.createCell(targetCol);
            MarkdownInline.setMarkdownRichTextCell(ctx.wb, c2, sp.lines.get(i), style);

            // ここだけ：indent を保持（0固定にしない）
            ctx.st.afterWriteNormalText(r2.getRowNum(), targetCol, indent, false);
        }

        // 末尾 <br> は次入力行へ継続（同列に縦展開）
        if (sp.endsWithBr) {
            ctx.st.pendingSameColBr = true;
            ctx.st.pendingSameColBrCol = targetCol;
            ctx.st.pendingSameColBrStyle = style;
            ctx.st.pendingSameColBrCarry = sp.carryPrefix;

            // <br> を跨いだあとに「同インデント連結」が暴発しないよう安全側で連結を切るなら：
            ctx.st.resetOnBlockBoundary(); // list context は消えない（RenderState実装のままなら）
        }
    }

    private static NormalTextFlags buildNormalTextFlags(int indent, RenderState st) {
        boolean isHeadingParagraph = st.inHeadingParagraphBlock && indent == 0 && !st.inListBlock;

        boolean isListNote = st.inListBlock && indent == 0 && st.lastRowType == RenderState.RowType.BLANK
                && st.lastBlankFromMarkdown && st.lastBlankRowIndex >= 0;

        boolean isListChildParagraph = indent > 0 && st.inListBlock && st.lastRowType == RenderState.RowType.BLANK
                && st.lastBlankFromMarkdown && st.lastBlankRowIndex >= 0;

        return new NormalTextFlags(isHeadingParagraph, isListNote, isListChildParagraph);
    }

    private static boolean shouldReuseBlankForNormalText(int indent, RenderState st, NormalTextFlags f) {
        if (st.lastBlankAfterTable) {
            return false;
        }
        boolean collapseEmptyLineBetweenPlainParagraphs = indent == 0 && st.lastRowType == RenderState.RowType.BLANK
                && st.lastBlankFromMarkdown && st.lastBlankRowIndex >= 0 && !st.inListBlock
                && (st.lastContentType == RenderState.ContentType.NORMAL
                        || st.lastContentType == RenderState.ContentType.CODE);

        return f.isListNote || f.isListChildParagraph || collapseEmptyLineBetweenPlainParagraphs;
    }

    private static int calcNormalTextCol(int indent, RenderState st, NormalTextFlags f) {
        if (f.isHeadingParagraph || f.isListNote) {
            return 0;
        }
        if (f.isListChildParagraph) {
            int parentDepth = ListStackUtil.getParentListDepthForChildParagraph(st.listStack);
            int col = 2 + Math.max(0, parentDepth); // C列起点
            return clampCol(col, st);
        }

        int baseCol;
        if (indent == 0) {
            baseCol = 0;
        } else if (!st.listStack.isEmpty()) {
            int depth = ListStackUtil.getDepthForIndent(st.listStack, indent);
            baseCol = 1 + depth;
        } else {
            int level = indent / 2;
            if (level < 0)
                level = 0;
            baseCol = 1 + level;
        }
        return clampCol(baseCol, st);
    }

    private static int clampCol(int col, RenderState st) {
        if (col < 0)
            return 0;
        if (col >= st.mergeLastCol)
            return st.mergeLastCol - 1;
        return col;
    }

    private static final class NormalTextFlags {
        final boolean isHeadingParagraph;
        final boolean isListNote;
        final boolean isListChildParagraph;

        NormalTextFlags(boolean hp, boolean ln, boolean lcp) {
            this.isHeadingParagraph = hp;
            this.isListNote = ln;
            this.isListChildParagraph = lcp;
        }
    }
}
