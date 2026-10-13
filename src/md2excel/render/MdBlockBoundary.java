package md2excel.render;

import java.util.EnumSet;

import md2excel.markdown.LineInfo;
import md2excel.markdown.LineKind;

final class MdBlockBoundary {

    private MdBlockBoundary() {
    }

    enum Action {
        CLOSE_TABLE, CLOSE_BLOCK_QUOTE, INSERT_AUTO_BLANK_IF_PREV_HEADING, // “番号付きは見出し直後に空行” 仕様
        RESET_PARAGRAPH, LEAVE_LIST_BLOCK_PRESERVING_LEVELS
    }

    // Policyに含まれる処理の実行順序。
    // 行を離れる際の表の終了は、processLineBoundary()でこれより先に判定する。
    private static final Action[] ORDER = { Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING, Action.RESET_PARAGRAPH,
            Action.LEAVE_LIST_BLOCK_PRESERVING_LEVELS };

    enum Policy {
        NONE(actions()),

        CODE_FENCE(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING, Action.RESET_PARAGRAPH)),

        MARKDOWN_BLANK(actions(Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH)),

        HORIZONTAL_RULE(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH, Action.LEAVE_LIST_BLOCK_PRESERVING_LEVELS)),

        HEADING(actions(Action.CLOSE_TABLE, Action.CLOSE_BLOCK_QUOTE, Action.RESET_PARAGRAPH, Action.LEAVE_LIST_BLOCK_PRESERVING_LEVELS)),

        BULLET_ITEM(actions(Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING)),

        NUMBER_ITEM(actions(Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING)),

        TABLE_LINE(actions(Action.CLOSE_BLOCK_QUOTE, Action.INSERT_AUTO_BLANK_IF_PREV_HEADING, Action.RESET_PARAGRAPH));

        final EnumSet<Action> actions;

        Policy(EnumSet<Action> actions) {
            this.actions = actions;
        }
    }

    /**
     * 保留段落の出力後、現在行を処理する前のブロック境界処理を行う。
     *
     * Policy自身が表を閉じる場合は、先行する表終了判定を省く。 それ以外の場合は、引用行や空行などへの遷移も含めて 現在の表を継続できるか判定する。
     */
    static void processLineBoundary(LineInfo line, RenderContext ctx) {
        Policy policy = policyFor(line);

        if (!policy.actions.contains(Action.CLOSE_TABLE)) {
            closeTableIfLeaving(line, ctx);
        }

        apply(policy, ctx);
    }

    /**
     * EOFで保留中の内容を確定する。
     *
     * 段落、コード枠、表、引用装飾の順序は従来どおり。 後続行用の自動空行挿入や通常のブロック境界リセットは行わない。
     */
    static void finishAtEof(ParagraphBuffer paragraph, RenderContext ctx) {
        if (paragraph != null) {
            ParagraphUtil.flush(paragraph, ctx);
        }

        if (ctx.st.codeBlock().isOpen()) {
            CodeBlockRenderer.finish(ctx);
        }

        if (ctx.st.isLastLineTable()) {
            MarkdownTableWriter.closeTableIfOpen(ctx.sheet, ctx.styles, ctx.st);
        }

        BlockQuoteUtil.closeBlockQuoteIfOpen(ctx.sheet, ctx.styles, ctx.st);
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

        for (Action x : a) {
            set.add(x);
        }

        return set;
    }

    static void apply(Policy p, RenderContext ctx) {
        EnumSet<Action> a = p.actions;

        for (Action act : ORDER) {
            if (!a.contains(act)) {
                continue;
            }

            switch (act) {
            case CLOSE_TABLE:
                MarkdownTableWriter.closeTableIfOpen(ctx.sheet, ctx.styles, ctx.st);
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

    static void closeTableIfLeaving(LineInfo nextLine, RenderContext ctx) {
        if (!ctx.st.isLastLineTable()) {
            return;
        }

        boolean continuesCurrentTable = nextLine != null && nextLine.isTableLike() && ctx.st.table().getQuoteDepth() == nextLine.getQuoteDepth();

        if (!continuesCurrentTable) {
            MarkdownTableWriter.closeTableIfOpen(ctx.sheet, ctx.styles, ctx.st);
        }
    }
}