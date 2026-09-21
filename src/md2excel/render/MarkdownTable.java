package md2excel.render;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import md2excel.excel.ExcelCellUtil;
import md2excel.excel.MdStyleCatalog;

public final class MarkdownTable {

    private MarkdownTable() {
    }

    public static boolean isTableLine(String line) {
        String trimmed = line.trim();
        return countUnescapedPipes(trimmed) >= 1;
    }

    public static boolean isTableSeparatorLine(String line) {
        if (line == null) {
            return false;
        }

        String trimmed = line.trim();
        if (countUnescapedPipes(trimmed) < 1) {
            return false;
        }

        List<String> cells = splitTableCells(trimmed);
        if (cells.isEmpty()) {
            return false;
        }

        for (String cell : cells) {
            if (!isTableSeparatorCell(cell)) {
                return false;
            }
        }

        return true;
    }

    public static boolean isTableStart(String headerLine, String separatorLine) {
        if (!isTableLine(headerLine) || isTableSeparatorLine(headerLine) || !isTableSeparatorLine(separatorLine)) {
            return false;
        }

        return splitTableCells(headerLine).size() == splitTableCells(separatorLine).size();
    }

    private static boolean isTableSeparatorCell(String cell) {
        StringBuilder compact = new StringBuilder();

        for (int i = 0; i < cell.length(); i++) {
            char ch = cell.charAt(i);
            if (!Character.isWhitespace(ch)) {
                compact.append(ch);
            }
        }

        int start = 0;
        int end = compact.length();

        if (start < end && compact.charAt(start) == ':') {
            start++;
        }
        if (start < end && compact.charAt(end - 1) == ':') {
            end--;
        }

        int hyphenCount = 0;
        for (int i = start; i < end; i++) {
            if (compact.charAt(i) != '-') {
                return false;
            }
            hyphenCount++;
        }

        // 既存の "| -- | -- |" も許容する
        return hyphenCount >= 1;
    }

    enum TableRowStyleRole {
        HEADER,
        BODY_WITH_BOTTOM_BORDER,
        BODY_WITHOUT_BOTTOM_BORDER
    }

    static final class TableRowRenderResult {
        final int firstRowNum;
        final int lastRowNum;
        final int lastCol;
        final List<TableRowStyleRole> rowStyleRoles;

        TableRowRenderResult(int firstRowNum, int lastRowNum, int lastCol, List<TableRowStyleRole> rowStyleRoles) {

            this.firstRowNum = firstRowNum;
            this.lastRowNum = lastRowNum;
            this.lastCol = lastCol;
            this.rowStyleRoles = new ArrayList<TableRowStyleRole>(rowStyleRoles);
        }

        TableRowStyleRole getStyleRole(int rowNum) {
            int index = rowNum - firstRowNum;

            if (index < 0 || index >= rowStyleRoles.size()) {
                throw new IllegalArgumentException("Row is outside table result: " + rowNum);
            }

            return rowStyleRoles.get(index);
        }
    }

    static TableRowRenderResult createTableRows(RenderContext ctx, String line, boolean isHeaderRow, int startCol) {
        List<String> rawCells = splitTableCells(line);
        if (!isHeaderRow && ctx.st.table().getEndCol() >= startCol) {

            int headerCellCount = ctx.st.table().getEndCol() - startCol + 1;

            if (rawCells.size() > headerCellCount) {
                rawCells = new ArrayList<String>(rawCells.subList(0, headerCellCount));
            } else {
                while (rawCells.size() < headerCellCount) {
                    rawCells.add("");
                }
            }
        }

        List<List<List<MarkdownInline.MdSegment>>> cellLines = new ArrayList<List<List<MarkdownInline.MdSegment>>>();
        int maxRowCount = 1;

        for (int i = 0; i < rawCells.size(); i++) {
            String colText = rawCells.get(i).trim();
            colText = unescapePipeOutsideInlineCode(colText);

            List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(colText);

            if (lines.isEmpty()) {
                lines = Collections.<List<MarkdownInline.MdSegment>>singletonList(
                        Collections.<MarkdownInline.MdSegment>emptyList());
            }

            cellLines.add(lines);

            if (lines.size() > maxRowCount) {
                maxRowCount = lines.size();
            }
        }

        int firstRowNum = -1;
        int lastRowNum = -1;
        int lastCol = startCol - 1;
        List<TableRowStyleRole> rowStyleRoles = new ArrayList<TableRowStyleRole>();

        for (int rowOffset = 0; rowOffset < maxRowCount; rowOffset++) {
            Row row = (rowOffset == 0) ? RowUtil.createRowOrReusePreviousMarkdownBlank(ctx.sheet, ctx.st,
                    RowUtil.ReuseKind.TABLE_ROW, ctx.styles.normalStyle)
                    : RowUtil.createRow(ctx.sheet, ctx.st, ctx.styles.normalStyle);

            boolean hasNextExpandedRow = rowOffset < maxRowCount - 1;

            TableRowStyleRole rowStyleRole;

            if (isHeaderRow) {
                rowStyleRole = TableRowStyleRole.HEADER;
            } else if (hasNextExpandedRow) {
                rowStyleRole = TableRowStyleRole.BODY_WITHOUT_BOTTOM_BORDER;
            } else {
                rowStyleRole = TableRowStyleRole.BODY_WITH_BOTTOM_BORDER;
            }

            rowStyleRoles.add(rowStyleRole);

            if (firstRowNum < 0) {
                firstRowNum = row.getRowNum();
            }
            lastRowNum = row.getRowNum();

            int colIndex = startCol;
            for (int c = 0; c < cellLines.size(); c++) {
                Cell cell = row.createCell(colIndex);

                List<List<MarkdownInline.MdSegment>> lines = cellLines.get(c);
                List<MarkdownInline.MdSegment> segments = (rowOffset < lines.size()) ? lines.get(rowOffset)
                        : Collections.<MarkdownInline.MdSegment>emptyList();

                if (isHeaderRow) {
                    if (!segments.isEmpty()) {
                        MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, segments,
                                ctx.styles.tableHeaderStyle);
                    } else {
                        cell.setCellStyle(ctx.styles.tableHeaderStyle);
                    }
                } else {
                    if (!segments.isEmpty()) {
                        MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, segments,
                                hasNextExpandedRow ? ctx.styles.tableBodyLastRowStyle : ctx.styles.tableBodyStyle);
                    } else {
                        cell.setCellStyle(
                                hasNextExpandedRow ? ctx.styles.tableBodyLastRowStyle : ctx.styles.tableBodyStyle);
                    }
                }

                colIndex++;
            }

            lastCol = Math.max(lastCol, colIndex - 1);
        }

        return new TableRowRenderResult(firstRowNum, lastRowNum, lastCol, rowStyleRoles);
    }

    private static List<String> splitTableCells(String line) {
        String trimmed = line.trim();
        String inner = trimmed;

        if (inner.startsWith("|")) {
            inner = inner.substring(1);
        }
        if (inner.endsWith("|")) {
            inner = inner.substring(0, inner.length() - 1);
        }

        List<String> cells = new ArrayList<String>();

        int segStart = 0;
        int n = inner.length();

        for (int i = 0; i <= n; i++) {
            if (i == n || (inner.charAt(i) == '|' && !isEscapedPipe(inner, i))) {
                cells.add(inner.substring(segStart, i));
                segStart = i + 1;
            }
        }

        return cells;
    }

    /**
     * pos の '|' が "\|" のようにエスケープされているか判定する。 直前に連続する '\' の個数が奇数ならエスケープ扱い。
     */
    private static boolean isEscapedPipe(String s, int pos) {
        if (pos <= 0 || pos >= s.length() || s.charAt(pos) != '|')
            return false;
        int bs = 0;
        for (int i = pos - 1; i >= 0 && s.charAt(i) == '\\'; i--) {
            bs++;
        }
        return (bs % 2) == 1;
    }

    private static int countUnescapedPipes(String s) {
        if (s == null || s.isEmpty())
            return 0;

        int count = 0;
        for (int i = 0; i < s.length(); i++) {
            if (s.charAt(i) == '|' && !isEscapedPipe(s, i)) {
                count++;
            }
        }
        return count;
    }

    /**
     * テーブルセル内の "\|" を "|" に戻す。
     */
    private static String unescapePipeOutsideInlineCode(String s) {
        if (s == null || s.isEmpty())
            return s;
        StringBuilder out = new StringBuilder(s.length());
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch == '\\' && i + 1 < s.length() && s.charAt(i + 1) == '|') {
                out.append('|');
                i++; // '|' を消費
                continue;
            }
            out.append(ch);
        }
        return out.toString();
    }

    public static void closeTableIfOpen(Sheet sheet, MdStyleCatalog styles, RenderState st) {

        if (!st.isLastLineTable()) {
            return;
        }

        TableState table = st.table();

        finalizeTableBorders(sheet, styles, table.getBodyStartRow(), table.getLastBodyRow(), table.getStartCol(),
                table.getEndCol());

        if (table.getLastBodyRow() >= 0) {
            st.updateBlockQuoteTableRowStyleRole(table.getLastBodyRow(), TableRowStyleRole.BODY_WITHOUT_BOTTOM_BORDER);
        }

        st.afterCloseTable();
    }

    private static void finalizeTableBorders(Sheet sheet, MdStyleCatalog styles, int bodyStartRow, int lastBodyRow,
            int startCol, int endCol) {

        if (lastBodyRow < 0 || bodyStartRow < 0)
            return;
        if (startCol < 0 || endCol < startCol)
            return;

        Row row = sheet.getRow(lastBodyRow);
        if (row == null)
            return;

        for (int c = startCol; c <= endCol; c++) {
            Cell cell = ExcelCellUtil.getOrCreateCell(row, c);
            cell.setCellStyle(styles.tableBodyLastRowStyle);
        }
    }

}