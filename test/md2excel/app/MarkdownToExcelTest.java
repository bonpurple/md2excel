package md2excel.app;

import static org.junit.Assert.assertEquals;

import java.io.IOException;

import org.junit.Test;

public class MarkdownToExcelTest {

    @Test
    public void failureMessageIncludesExceptionMessage() {
        String message = MarkdownToExcel.createFailureMessage(new IOException("出力ファイルを開けません。"));

        assertEquals("Excelファイルの生成に失敗しました。\n\n" + "出力ファイルを開けません。", message);
    }

    @Test
    public void failureMessageUsesClassNameWhenMessageIsNull() {
        String message = MarkdownToExcel.createFailureMessage(new IOException());

        assertEquals("Excelファイルの生成に失敗しました。\n\n" + "IOException", message);
    }
}