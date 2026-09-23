package md2excel.render;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import md2excel.excel.ExcelCellUtil;
import md2excel.excel.MdStyleCatalog;

public final class MarkdownTable {

    private MarkdownTable() {
    }

    public static boolean isTableLine(String line) {
        return MarkdownTableParser.isTableLine(line);
    }

    public static boolean isTableSeparatorLine(String line) {
        return MarkdownTableParser.isTableSeparatorLine(line);
    }

    public static boolean isTableStart(String headerLine, String separatorLine) {
        return MarkdownTableParser.isTableStart(headerLine, separatorLine);
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
        private final List<TableRowStyleRole> rowStyleRoles;

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
        List<String> rawCells = MarkdownTableParser.splitTableCells(line);
        rawCells = normalizeCellCount(ctx, rawCells, isHeaderRow, startCol);
        List<List<List<MarkdownInline.MdSegment>>> cellLines = parseCellDisplayLines(rawCells);
        int maxRowCount = 1;

        for (List<List<MarkdownInline.MdSegment>> lines : cellLines) {
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
            CellStyle rowStyle = getTableCellStyle(ctx.styles, rowStyleRole);

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

                if (!segments.isEmpty()) {
                    MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell, segments, rowStyle);
                } else {
                    cell.setCellStyle(rowStyle);
                }

                colIndex++;
            }

            lastCol = Math.max(lastCol, colIndex - 1);
        }

        return new TableRowRenderResult(firstRowNum, lastRowNum, lastCol, rowStyleRoles);
    }

    private static List<String> normalizeCellCount(RenderContext ctx, List<String> rawCells, boolean isHeaderRow,
            int startCol) {
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

        return rawCells;
    }

    private static List<List<List<MarkdownInline.MdSegment>>> parseCellDisplayLines(List<String> rawCells) {
        List<List<List<MarkdownInline.MdSegment>>> cellLines = new ArrayList<List<List<MarkdownInline.MdSegment>>>();

        for (int i = 0; i < rawCells.size(); i++) {
            String colText = rawCells.get(i).trim();
            colText = MarkdownTableParser.unescapePipeOutsideInlineCode(colText);

            List<List<MarkdownInline.MdSegment>> lines = MarkdownInline.parseParagraphToDisplayLines(colText);

            if (lines.isEmpty()) {
                lines = Collections.<List<MarkdownInline.MdSegment>>singletonList(
                        Collections.<MarkdownInline.MdSegment>emptyList());
            }

            cellLines.add(lines);
        }

        return cellLines;
    }

    private static CellStyle getTableCellStyle(MdStyleCatalog styles, TableRowStyleRole role) {
        switch (role) {
        case HEADER:
            return styles.tableHeaderStyle;
        case BODY_WITHOUT_BOTTOM_BORDER:
            return styles.tableBodyLastRowStyle;
        case BODY_WITH_BOTTOM_BORDER:
            return styles.tableBodyStyle;
        default:
            throw new IllegalArgumentException("Unknown table row style role: " + role);
        }
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
