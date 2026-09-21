package md2excel.render;

import java.util.EnumSet;

import md2excel.markdown.LineInfo;
import md2excel.markdown.LineKind;

public final class MdBlockBoundary {
    private MdBlockBoundary() {
    }

    enum Action {
        CLOSE_TABLE,
        CLOSE_BLOCK_QUOTE,
        INSERT_AUTO_BLANK_IF_PREV_HEADING, // “番号付きは見出し直後に空行” 仕様
        RESET_PARAGRAPH,
        LEAVE_LIST_BLOCK_PRESERVING_LEVELS
    }

    // ここが「ブロック境界の order の唯一の定義」
    private static final Action[] ORDER = { Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE,
            Action.INSERT_AUTO_BLANK_IF_PREV_HEADING, Action.RESET_PARAGRAPH,
            Action.LEAVE_LIST_BLOCK_PRESERVING_LEVELS };

    public enum Policy {
        NONE(actions()),

        CODE_FENCE(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING,
                Action.RESET_PARAGRAPH)),

        MARKDOWN_BLANK(actions(Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH)),

        HORIZONTAL_RULE(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH,
                Action.LEAVE_LIST_BLOCK_PRESERVING_LEVELS)),

        HEADING(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH,
                Action.LEAVE_LIST_BLOCK_PRESERVING_LEVELS)),

        BULLET_ITEM(actions(Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING)),

        NUMBER_ITEM(actions(Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING)),

        TABLE_LINE(actions(Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING, Action.RESET_PARAGRAPH));

        final EnumSet<Action> actions;

        Policy(EnumSet<Action> actions) {
            this.actions = actions;
        }
    }

    static Policy policyFor(LineKind kind) {
        if (kind == null) {
            throw new IllegalArgumentException("kind must not be null");
        }

        switch (kind) {
        case CODE_FENCE:
            return Policy.CODE_FENCE;

        case CODE_LINE:
            return Policy.NONE;

        case BLANK:
            return Policy.MARKDOWN_BLANK;

        case HORIZONTAL_RULE:
            return Policy.HORIZONTAL_RULE;

        case TABLE_SEPARATOR:
        case TABLE_ROW:
            return Policy.TABLE_LINE;

        case HEADING:
            return Policy.HEADING;

        case BULLET_ITEM:
            return Policy.BULLET_ITEM;

        case NUMBER_ITEM:
            return Policy.NUMBER_ITEM;

        case NORMAL:
            return Policy.NONE;

        default:
            throw new AssertionError("Unhandled LineKind: " + kind);
        }
    }

    static Policy policyFor(LineInfo line) {
        if (line == null) {
            throw new IllegalArgumentException("line must not be null");
        }

        // 引用の内側の種類に関係なく、
        // 引用ブロックとしてMarkdownRenderer側で処理する。
        if (line.isQuoted()) {
            return Policy.NONE;
        }

        return policyFor(line.getKind());
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
                ctx.st.ensureAutoBlankIfPrevHeading(ctx.sheet, ctx.styles.blankRowStyle);
                break;

            case RESET_PARAGRAPH:
                ctx.st.resetOnBlockBoundary();
                break;

            case LEAVE_LIST_BLOCK_PRESERVING_LEVELS:
                ctx.st.leaveListBlockPreservingLevels();
                break;
            }
        }
    }

    public static void closeTableIfLeaving(LineInfo nextLine, RenderContext ctx) {

        if (!ctx.st.isLastLineTable()) {
            return;
        }

        boolean continuesCurrentTable = nextLine != null && nextLine.isTableLike()
                && ctx.st.table().getQuoteDepth() == nextLine.getQuoteDepth();

        if (!continuesCurrentTable) {
            MarkdownTable.closeTableIfOpen(ctx.sheet, ctx.styles, ctx.st);
        }
    }
}