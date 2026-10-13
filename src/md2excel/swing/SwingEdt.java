package md2excel.swing;

import java.lang.reflect.InvocationTargetException;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Supplier;

import javax.swing.SwingUtilities;

/**
 * 処理をSwingのEDT上で同期実行する。
 */
public final class SwingEdt {

    private SwingEdt() {
    }

    public static <T> T call(Supplier<T> task, String interruptedMessage, String failedMessage) {

        if (SwingUtilities.isEventDispatchThread()) {
            return task.get();
        }

        AtomicReference<T> result = new AtomicReference<T>();

        try {
            SwingUtilities.invokeAndWait(() -> result.set(task.get()));

        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IllegalStateException(interruptedMessage, e);

        } catch (InvocationTargetException e) {
            Throwable cause = e.getCause();

            if (cause instanceof RuntimeException) {
                throw (RuntimeException)cause;
            }

            if (cause instanceof Error) {
                throw (Error)cause;
            }

            throw new IllegalStateException(failedMessage, cause);
        }

        return result.get();
    }

    public static void run(Runnable task, String interruptedMessage, String failedMessage) {

        call(() -> {
            task.run();
            return null;
        }, interruptedMessage, failedMessage);
    }
}