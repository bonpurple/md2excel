package md2excel.render;

import java.util.Iterator;

/**
 * 入力を1行先読みするカーソル。
 */
final class LineCursor {

    private final Iterator<String> source;

    private boolean hasBuffered;
    private String buffered;

    LineCursor(Iterator<String> source) {
        this.source = source;
        advance();
    }

    boolean hasNext() {
        return hasBuffered;
    }

    String next() {
        String current = buffered;
        advance();
        return current;
    }

    String peek() {
        return hasBuffered ? buffered : null;
    }

    private void advance() {
        hasBuffered = source.hasNext();
        buffered = hasBuffered ? source.next() : null;
    }
}