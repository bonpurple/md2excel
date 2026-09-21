package md2excel.render;

import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import md2excel.excel.ExcelCellUtil;

public final class CellAppendUtil {

    private CellAppendUtil() {
    }

    public static void appendMarkdownWithSpace(RenderContext ctx, int rowNum, int colNum, String markdownText,
            CellStyle baseStyle) {

        appendMarkdown(ctx, rowNum, colNum, markdownText, baseStyle, true);
    }

    public static void appendMarkdown(RenderContext ctx, int rowNum, int colNum, String markdownText,
            CellStyle baseStyle, boolean withLeadingSpace) {

        if (markdownText == null || markdownText.isEmpty()) {
            return;
        }

        List<MarkdownInline.MdSegment> segments = MarkdownInline.parseParagraphToSingleLineSegments(markdownText);

        appendSegments(ctx, rowNum, colNum, segments, baseStyle, withLeadingSpace, true);
    }

    public static void appendResolvedSegmentsWithSpace(RenderContext ctx, int rowNum, int colNum,
            List<MarkdownInline.MdSegment> segments, CellStyle baseStyle) {

        appendResolvedSegments(ctx, rowNum, colNum, segments, baseStyle, true);
    }

    public static void appendResolvedSegments(RenderContext ctx, int rowNum, int colNum,
            List<MarkdownInline.MdSegment> segments, CellStyle baseStyle, boolean withLeadingSpace) {

        appendSegments(ctx, rowNum, colNum, segments, baseStyle, withLeadingSpace, false);
    }

    private static void appendSegments(RenderContext ctx, int rowNum, int colNum,
            List<MarkdownInline.MdSegment> segments, CellStyle baseStyle, boolean withLeadingSpace,
            boolean initializeWhenEmpty) {

        boolean empty = segments == null || segments.isEmpty();

        if (empty && !initializeWhenEmpty) {
            return;
        }

        Row row = RowUtil.getOrCreateRow(ctx.sheet, rowNum, ctx.styles.normalStyle);

        Cell cell = row.getCell(colNum);

        if (cell == null) {
            cell = ExcelCellUtil.getOrCreateCell(row, colNum);

            MarkdownInline.setResolvedSegmentsCell(ctx.fontCache, cell,
                    Collections.<MarkdownInline.MdSegment>emptyList(), baseStyle);
        }

        if (!empty) {
            MarkdownInline.appendResolvedSegmentsToCell(ctx.fontCache, cell, segments, baseStyle, withLeadingSpace);
        }
    }
}