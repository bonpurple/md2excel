package md2excel;

import java.util.List;

public final class ListStackUtil {

    private ListStackUtil() {
    }

    public static class ListLevel {
        public int indent;
        public boolean ordered;

        public ListLevel(int indent, boolean ordered) {
            this.indent = indent;
            this.ordered = ordered;
        }
    }

    public static int updateListDepth(List<ListLevel> listStack, int indent, boolean ordered) {
        if (listStack.isEmpty()) {
            listStack.add(new ListLevel(indent, ordered));
            return 0;
        }

        int depth = listStack.size() - 1;
        ListLevel last = listStack.get(depth);

        if (indent > last.indent) {
            listStack.add(new ListLevel(indent, ordered));
            return listStack.size() - 1;
        }

        while (depth > 0 && indent < listStack.get(depth).indent) {
            listStack.remove(depth);
            depth--;
        }

        ListLevel cur = listStack.get(depth);
        cur.indent = indent;
        cur.ordered = ordered;

        return depth;
    }

    public static int getDepthForIndent(List<ListLevel> stack, int indentSpaces) {
        if (indentSpaces < 0)
            indentSpaces = 0;
        if (stack.isEmpty())
            return 0;

        int depth = 0;
        for (int i = 0; i < stack.size(); i++) {
            if (indentSpaces > stack.get(i).indent) {
                depth = i + 1;
            } else {
                break;
            }
        }
        return depth;
    }

    public static int getParentListDepthForChildParagraph(List<ListLevel> listStack) {
        if (listStack == null || listStack.isEmpty()) {
            return 0;
        }
        for (int i = listStack.size() - 1; i >= 0; i--) {
            if (listStack.get(i).ordered) {
                return i;
            }
        }
        return listStack.size() - 1;
    }
}
