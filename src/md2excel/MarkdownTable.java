package md2excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public final class MarkdownTable {

    private MarkdownTable() {
    }

    public static boolean isTableLine(String line) {
        String trimmed = line.trim();
        if (!trimmed.startsWith("|"))
            return false;
        int first = trimmed.indexOf('|');
        int last = trimmed.lastIndexOf('|');
        return first != -1 && last != -1 && first != last;
    }

    public static boolean isTableSeparatorLine(String trimmed) {
        if (!trimmed.contains("|"))
            return false;
        for (int i = 0; i < trimmed.length(); i++) {
            char c = trimmed.charAt(i);
            if (c != '|' && c != '-' && c != ':' && !Character.isWhitespace(c)) {
                return false;
            }
        }
        return true;
    }

    public static int createTableRow(Workbook wb, String line, Row row, MdStyle styles, boolean isHeaderRow,
            int startCol) {

        String trimmed = line.trim();
        if (!trimmed.startsWith("|")) {
            return startCol - 1;
        }

        // 既存実装と同じく「先頭と末尾の1文字」を落とす（末尾が '|' である前提の仕様）
        String inner = trimmed.substring(1, trimmed.length() - 1);

        int colIndex = startCol;

        int segStart = 0;
        int n = inner.length();
        boolean inCode = false; // `...` 内は | を区切りにしない（安全側）
        for (int i = 0; i <= n; i++) {
            if (i < n && inner.charAt(i) == '`') {
                inCode = !inCode;
                continue;
            }
            if (i == n || (inner.charAt(i) == '|' && !inCode && !isEscapedPipe(inner, i))) {
                String colText = inner.substring(segStart, i).trim();

                // \| を | として扱う（インラインコード内も）
                colText = unescapePipeOutsideInlineCode(colText);

                // <br> を空白へ（インラインコード内は触らない）
                colText = MdTextUtil.replaceBrOutsideInlineCode(colText, " ");
                colText = MdTextUtil.collapseSpaces(colText);

                Cell cell = row.createCell(colIndex++);

                if (isHeaderRow) {
                    String joined = MarkdownInline.brToSingleSpace(colText); // ★必ず split を通る
                    if (!joined.isEmpty()) {
                        MarkdownInline.setMarkdownRichTextCell(wb, cell, joined, styles.tableHeaderStyle);
                    } else {
                        cell.setCellStyle(styles.tableHeaderStyle);
                    }
                } else {
                    String joined = MarkdownInline.brToSingleSpace(colText); // ★必ず split を通る
                    if (!joined.isEmpty()) {
                        MarkdownInline.setMarkdownRichTextCell(wb, cell, joined, styles.tableBodyStyle);
                    } else {
                        cell.setCellStyle(styles.tableBodyStyle);
                    }
                }

                segStart = i + 1; // 次のセグメント開始
            }
        }

        return colIndex - 1;
    }
    
    /**
     * pos の '|' が "\|" のようにエスケープされているか判定する。
     * 直前に連続する '\' の個数が奇数ならエスケープ扱い。
     */
    private static boolean isEscapedPipe(String s, int pos) {
        if (pos <= 0 || pos >= s.length() || s.charAt(pos) != '|') return false;
        int bs = 0;
        for (int i = pos - 1; i >= 0 && s.charAt(i) == '\\'; i--) {
            bs++;
        }
        return (bs % 2) == 1;
    }

    /**
     * テーブルセル内の "\|" を "|" に戻す。
     */
    private static String unescapePipeOutsideInlineCode(String s) {
        if (s == null || s.isEmpty()) return s;
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

    public static void closeTableIfOpen(org.apache.poi.ss.usermodel.Sheet sheet, MdStyle styles, RenderState st) {
        if (!st.lastLineWasTable)
            return;

        finalizeTableBorders(sheet, styles, st.currentTableHeaderRow, st.currentTableBodyStartRow,
                st.currentTableLastBodyRow, st.currentTableStartCol, st.currentTableEndCol);

        st.lastLineWasTable = false;
        st.currentTableHeaderRow = -1;
        st.currentTableBodyStartRow = -1;
        st.currentTableLastBodyRow = -1;
        st.currentTableStartCol = 0;
        st.currentTableEndCol = -1;
    }

    private static void finalizeTableBorders(org.apache.poi.ss.usermodel.Sheet sheet, MdStyle styles, int headerRow,
            int bodyStartRow, int lastBodyRow, int startCol, int endCol) {

        if (lastBodyRow < 0 || bodyStartRow < 0)
            return;
        if (startCol < 0 || endCol < startCol)
            return;

        Row row = sheet.getRow(lastBodyRow);
        if (row == null)
            return;

        for (int c = startCol; c <= endCol; c++) {
            Cell cell = row.getCell(c);
            if (cell == null)
                cell = row.createCell(c);
            cell.setCellStyle(styles.tableBodyLastRowStyle);
        }
    }

}
