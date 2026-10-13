package md2excel.render;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.util.Arrays;
import java.util.Collections;

import org.junit.Test;

public class LineCursorTest {

    @Test
    public void peekReturnsNextLineAfterEachRead() {
        LineCursor cursor = new LineCursor(Arrays.asList("header", "separator", "body").iterator());

        assertTrue(cursor.hasNext());
        assertEquals("header", cursor.peek());

        assertEquals("header", cursor.next());
        assertEquals("separator", cursor.peek());

        assertEquals("separator", cursor.next());
        assertEquals("body", cursor.peek());

        assertEquals("body", cursor.next());
        assertFalse(cursor.hasNext());
        assertNull(cursor.peek());
    }

    @Test
    public void emptyInputHasNoBufferedLine() {
        LineCursor cursor = new LineCursor(Collections.<String> emptyList().iterator());

        assertFalse(cursor.hasNext());
        assertNull(cursor.peek());
    }
}