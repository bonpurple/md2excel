package md2excel.render;

final class RenderLayout {

    private RenderLayout() {
    }

    static int clampCol(int col, RenderState st) {
        if (col < 0) {
            return 0;
        }

        if (col >= st.getRenderEndColExclusive()) {
            return st.getRenderEndColExclusive() - 1;
        }

        return col;
    }

    static int rootCol(RenderState st) {
        return clampCol(st.getStartColIndex(), st);
    }

    static int calcBlockStartCol(int indent, RenderState st) {
        if (indent <= 0) {
            return rootCol(st);
        }

        int col;

        if (st.hasListLevels()) {
            int depth = st.getListDepthForIndent(indent);

            col = st.getStartColIndex() + 1 + depth;
        } else {
            int level = Math.max(0, indent / 2);
            col = st.getStartColIndex() + 1 + level;
        }

        return clampCol(col, st);
    }

    static int calcQuoteStartCol(int indent, RenderState st) {
        return clampCol(calcBlockStartCol(indent, st) + 1, st);
    }

    static int calcQuoteContentCol(int quoteStartCol, int quoteDepth, RenderState st) {

        return clampCol(quoteStartCol + Math.max(1, quoteDepth) - 1, st);
    }
}