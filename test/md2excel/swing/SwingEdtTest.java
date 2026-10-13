package md2excel.swing;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertSame;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import javax.swing.SwingUtilities;

import org.junit.Test;

public class SwingEdtTest {

    private static final String INTERRUPTED = "interrupted";
    private static final String FAILED = "failed";

    @Test
    public void callFromBackgroundRunsOnEdtAndReturnsValue() {
        String result = SwingEdt.call(() -> {
            assertTrue(SwingUtilities.isEventDispatchThread());
            return "result";
        }, INTERRUPTED, FAILED);

        assertEquals("result", result);
    }

    @Test
    public void callFromEdtRunsDirectlyAndReturnsValue() throws Exception {
        AtomicReference<String> result = new AtomicReference<String>();

        SwingUtilities.invokeAndWait(() -> result.set(SwingEdt.call(() -> {
            assertTrue(SwingUtilities.isEventDispatchThread());
            return "result";
        }, INTERRUPTED, FAILED)));

        assertEquals("result", result.get());
    }

    @Test
    public void runExecutesOnEdt() {
        AtomicBoolean called = new AtomicBoolean();

        SwingEdt.run(() -> {
            assertTrue(SwingUtilities.isEventDispatchThread());
            called.set(true);
        }, INTERRUPTED, FAILED);

        assertTrue(called.get());
    }

    @Test
    public void propagatesOriginalRuntimeException() {
        RuntimeException failure = new IllegalArgumentException("test");

        try {
            SwingEdt.run(() -> {
                throw failure;
            }, INTERRUPTED, FAILED);
            fail("RuntimeException was expected");

        } catch (RuntimeException actual) {
            assertSame(failure, actual);
        }
    }

    @Test
    public void propagatesOriginalError() {
        Error failure = new AssertionError("test");

        try {
            SwingEdt.run(() -> {
                throw failure;
            }, INTERRUPTED, FAILED);
            fail("Error was expected");

        } catch (Error actual) {
            assertSame(failure, actual);
        }
    }
}