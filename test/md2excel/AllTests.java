package md2excel;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;

import md2excel.app.BlockBoundaryRenderingTest;
import md2excel.app.MarkdownFileSourceTest;
import md2excel.app.MarkdownToExcelConverterTest;
import md2excel.app.MarkdownToExcelTest;
import md2excel.app.MarkdownWorkbookRendererTest;
import md2excel.app.ParagraphRenderingTest;
import md2excel.app.TableRenderingTest;
import md2excel.app.SavedFormattingTest;
import md2excel.config.Md2ExcelConfigDialogTest;
import md2excel.config.Md2ExcelConfigTest;
import md2excel.config.MdFontSettingsTest;
import md2excel.config.MdSheetSettingsTest;
import md2excel.excel.CodeBlockFrameMaskTest;
import md2excel.excel.MdStyleCatalogTest;
import md2excel.markdown.CodeFenceTest;
import md2excel.markdown.ListStackUtilTest;
import md2excel.markdown.MdCharUtilTest;
import md2excel.markdown.MdInlineCodeUtilTest;
import md2excel.markdown.MdTextUtilTest;
import md2excel.markdown.NumberedListMarkerTest;
import md2excel.render.ListRenderStateTest;
import md2excel.render.MarkdownFontCacheTest;
import md2excel.render.MarkdownInlineTest;
import md2excel.render.MarkdownLineParserTest;
import md2excel.render.MarkdownTableTest;
import md2excel.render.SheetColumnLayoutTest;

@RunWith(Suite.class)
@Suite.SuiteClasses({ MdCharUtilTest.class, CodeFenceTest.class, NumberedListMarkerTest.class, MdTextUtilTest.class,
        MdInlineCodeUtilTest.class, ListStackUtilTest.class, ListRenderStateTest.class, MdFontSettingsTest.class,
        MdSheetSettingsTest.class, Md2ExcelConfigTest.class, Md2ExcelConfigDialogTest.class,
        CodeBlockFrameMaskTest.class, MdStyleCatalogTest.class, MarkdownInlineTest.class, MarkdownFontCacheTest.class,
        SheetColumnLayoutTest.class, MarkdownTableTest.class, MarkdownLineParserTest.class,
        MarkdownFileSourceTest.class, MarkdownWorkbookRendererTest.class, MarkdownToExcelTest.class,
        MarkdownToExcelConverterTest.class, ParagraphRenderingTest.class, BlockBoundaryRenderingTest.class,
        TableRenderingTest.class, SavedFormattingTest.class })
public class AllTests {
}
