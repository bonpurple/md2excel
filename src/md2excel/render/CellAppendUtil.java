package md2excel.render;

import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import md2excel.excel.MdStyle;

public final class CellAppendUtil {
    private CellAppendUtil() {
    }

    public static void appendMarkdownWithSpace(Sheet sheet, Workbook wb, MdStyle styles, int rowNum, int colNum,
            String markdownText, CellStyle baseStyle) {
        appendMarkdown(sheet, wb, styles, rowNum, colNum, markdownText, baseStyle, true);
    }

    public static void appendMarkdown(Sheet sheet, Workbook wb, MdStyle styles, int rowNum, int colNum,
            String markdownText, CellStyle baseStyle, boolean withLeadingSpace) {

        if (markdownText == null || markdownText.isEmpty()) {
            return;
        }

        Row row = RowUtil.getOrCreateRow(sheet, rowNum, styles.normalStyle);

        Cell cell = row.getCell(colNum);
        if (cell == null) {
            cell = row.createCell(colNum);
            MarkdownInline.setResolvedSegmentsCell(wb, cell, Collections.<MarkdownInline.MdSegment>emptyList(),
                    baseStyle);
        }

        MarkdownInline.appendResolvedSegmentsToCell(wb, cell,
                MarkdownInline.splitByBrPreserveFormatting(markdownText).lines.isEmpty()
                        ? Collections.<MarkdownInline.MdSegment>emptyList()
                        : MarkdownInline
                                .joinLinesWithSingleSpace(MarkdownInline.splitByBrPreserveFormatting(markdownText)),
                baseStyle, withLeadingSpace);
    }

    // resolved segment
    public static void appendResolvedSegmentsWithSpace(Sheet sheet, Workbook wb, MdStyle styles, int rowNum, int colNum,
            List<MarkdownInline.MdSegment> segments, CellStyle baseStyle) {
        appendResolvedSegments(sheet, wb, styles, rowNum, colNum, segments, baseStyle, true);
    }

    public static void appendResolvedSegments(Sheet sheet, Workbook wb, MdStyle styles, int rowNum, int colNum,
            List<MarkdownInline.MdSegment> segments, CellStyle baseStyle, boolean withLeadingSpace) {

        if (segments == null || segments.isEmpty()) {
            return;
        }

        Row row = RowUtil.getOrCreateRow(sheet, rowNum, styles.normalStyle);

        Cell cell = row.getCell(colNum);
        if (cell == null) {
            cell = row.createCell(colNum);
            MarkdownInline.setResolvedSegmentsCell(wb, cell, Collections.<MarkdownInline.MdSegment>emptyList(),
                    baseStyle);
        }

        MarkdownInline.appendResolvedSegmentsToCell(wb, cell, segments, baseStyle, withLeadingSpace);
    }

    public static void appendMarkdownWithSpace(RenderContext ctx, int rowNum, int colNum, String markdownText,
            CellStyle baseStyle) {
        appendMarkdown(ctx, rowNum, colNum, markdownText, baseStyle, true);
    }

    public static void appendMarkdown(RenderContext ctx, int rowNum, int colNum, String markdownText,
            CellStyle baseStyle, boolean withLeadingSpace) {
        appendMarkdown(ctx.sheet, ctx.wb, ctx.styles, rowNum, colNum, markdownText, baseStyle, withLeadingSpace);
    }

    // ctx版
    public static void appendResolvedSegmentsWithSpace(RenderContext ctx, int rowNum, int colNum,
            List<MarkdownInline.MdSegment> segments, CellStyle baseStyle) {
        appendResolvedSegments(ctx.sheet, ctx.wb, ctx.styles, rowNum, colNum, segments, baseStyle, true);
    }

    public static void appendResolvedSegments(RenderContext ctx, int rowNum, int colNum,
            List<MarkdownInline.MdSegment> segments, CellStyle baseStyle, boolean withLeadingSpace) {
        appendResolvedSegments(ctx.sheet, ctx.wb, ctx.styles, rowNum, colNum, segments, baseStyle, withLeadingSpace);
    }
}