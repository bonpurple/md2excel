package md2excel.app;

import java.io.IOException;
import java.nio.file.Path;

import javax.swing.JOptionPane;

import md2excel.config.Md2ExcelConfig;
import md2excel.config.Md2ExcelConfigDialog;
import md2excel.swing.SwingEdt;

public final class MarkdownToExcel {

    private static final int EXIT_CODE_ERROR = 1;

    private MarkdownToExcel() {
    }

    public static void main(String[] args) {
        try {
            run();

        } catch (Exception e) {
            reportFailure(e);
            System.exit(EXIT_CODE_ERROR);
        }
    }

    /**
     * 設定の読み込み、変換、完了通知を実行する。
     *
     * キャンセルされた場合は何も生成せず正常終了する。
     */
    private static void run() throws IOException {
        Md2ExcelConfig config = Md2ExcelConfigDialog.load();

        if (config == null) {
            System.out.println("キャンセルされました。処理を終了します。");
            return;
        }

        MarkdownToExcelConverter converter = new MarkdownToExcelConverter();

        converter.convert(config);

        Path outputPath = config.getOutputPath().toAbsolutePath();

        System.out.println("生成完了: " + outputPath);

        showMessageDialogOnEdt("Excel ファイルを生成しました。\n" + outputPath, "完了", JOptionPane.INFORMATION_MESSAGE);
    }

    /**
     * 例外の詳細を標準エラーへ出力し、 ユーザー向けのエラーダイアログを表示する。
     */
    private static void reportFailure(Exception failure) {
        System.err.println("Excelファイルの生成に失敗しました。");

        failure.printStackTrace(System.err);

        String message = createFailureMessage(failure);

        try {
            showMessageDialogOnEdt(message, "エラー", JOptionPane.ERROR_MESSAGE);

        } catch (RuntimeException dialogFailure) {
            /*
             * エラーダイアログの表示失敗で、 本来の例外情報を失わないようにする。
             */
            failure.addSuppressed(dialogFailure);

            System.err.println("エラーダイアログの表示にも失敗しました。");

            dialogFailure.printStackTrace(System.err);
        }
    }

    static String createFailureMessage(Exception failure) {

        String detail = failure.getMessage();

        if (detail == null || detail.trim().isEmpty()) {
            detail = failure.getClass().getSimpleName();
        }

        return "Excelファイルの生成に失敗しました。\n\n" + detail;
    }

    /**
     * JOptionPaneをSwingのEDT上で表示する。
     */
    private static void showMessageDialogOnEdt(final String message, final String title, final int messageType) {

        SwingEdt.run(() -> JOptionPane.showMessageDialog(null, message, title, messageType), "ダイアログの表示中に処理が中断されました。", "ダイアログの表示に失敗しました。");
    }
}