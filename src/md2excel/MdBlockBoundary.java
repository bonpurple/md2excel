package md2excel;

import java.util.EnumSet;

public final class MdBlockBoundary {
    private MdBlockBoundary() {
    }

    enum Action {
        CLOSE_TABLE,
        CLOSE_BLOCK_QUOTE,
        INSERT_AUTO_BLANK_IF_PREV_HEADING, // “番号付きは見出し直後に空行” 仕様
        RESET_PARAGRAPH,
        CLEAR_LIST_CONTEXT
    }

    // ここが「ブロック境界の order の唯一の定義」
    private static final Action[] ORDER = { Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE,
            Action.INSERT_AUTO_BLANK_IF_PREV_HEADING, Action.RESET_PARAGRAPH, Action.CLEAR_LIST_CONTEXT };

    public enum Policy {
        NONE(actions()),
        CODE_FENCE(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH)),
        MARKDOWN_BLANK(actions(Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH)),
        HORIZONTAL_RULE(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH,
                Action.CLEAR_LIST_CONTEXT)),
        HEADING(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH,
                Action.CLEAR_LIST_CONTEXT)),
        BULLET_ITEM(actions(Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH)),
        NUMBER_ITEM(actions(Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING,
                Action.RESET_PARAGRAPH)),
        TABLE_LINE(actions(Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH));

        final EnumSet<Action> actions;

        Policy(EnumSet<Action> actions) {
            this.actions = actions;
        }
    }

    static EnumSet<Action> actions(Action... a) {
        EnumSet<Action> set = EnumSet.noneOf(Action.class);
        for (Action x : a)
            set.add(x);
        return set;
    }

    public static void apply(Policy p, RenderContext ctx) {
        EnumSet<Action> a = p.actions;

        for (Action act : ORDER) {
            if (!a.contains(act))
                continue;

            switch (act) {
            case CLOSE_TABLE:
                MarkdownTable.closeTableIfOpen(ctx.sheet, ctx.styles, ctx.st);
                break;

            case CLOSE_BLOCK_QUOTE:
                BlockQuoteUtil.closeBlockQuoteIfOpen(ctx.sheet, ctx.styles, ctx.st);
                break;

            case INSERT_AUTO_BLANK_IF_PREV_HEADING:
                ctx.st.ensureAutoBlankIfPrevHeading(ctx.sheet, ctx.styles.normalStyle);
                break;

            case RESET_PARAGRAPH:
                ctx.st.resetOnBlockBoundary();
                break;

            case CLEAR_LIST_CONTEXT:
                ctx.st.clearListContext();
                break;
            }
        }
    }

    public static void closeTableIfLeaving(boolean tableLine, RenderContext ctx) {
        if (ctx.st.lastLineWasTable && !tableLine) {
            MarkdownTable.closeTableIfOpen(ctx.sheet, ctx.styles, ctx.st);
        }
    }
}
