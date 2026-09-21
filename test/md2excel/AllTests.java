package md2excel;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;

import md2excel.app.MarkdownToExcelConverterTest;
import md2excel.config.Md2ExcelConfigTest;
import md2excel.markdown.ListStackUtilTest;
import md2excel.markdown.MdInlineCodeUtilTest;
import md2excel.markdown.MdTextUtilTest;
import md2excel.render.MarkdownLineParserTest;
import md2excel.render.MarkdownTableTest;

@RunWith(Suite.class)
@Suite.SuiteClasses({ MdTextUtilTest.class, MdInlineCodeUtilTest.class, ListStackUtilTest.class,
        Md2ExcelConfigTest.class, MarkdownTableTest.class, MarkdownLineParserTest.class,
        MarkdownToExcelConverterTest.class })
public class AllTests {
}