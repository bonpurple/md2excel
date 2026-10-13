package md2excel.render;

/**
 * 解析済みインラインテキストの1区間。
 */
final class MdSegment {

    final String text;
    final boolean inBold;
    final boolean inItalic;
    final boolean inCode;

    MdSegment(String text, boolean inBold, boolean inItalic, boolean inCode) {
        this.text = text;
        this.inBold = inBold;
        this.inItalic = inItalic;
        this.inCode = inCode;
    }
}