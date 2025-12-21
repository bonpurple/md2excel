package md2excel;

public final class MdTextUtil {
    private MdTextUtil() {
    }

    public static int countLeadingSpacesOrTabs(String s) {
        int count = 0;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch == ' ')
                count++;
            else if (ch == '\t')
                count += 4;
            else
                break;
        }
        return count;
    }

    public static int countHeadingLevel(String trimmedLine) {
        int count = 0;
        for (int i = 0; i < trimmedLine.length(); i++) {
            if (trimmedLine.charAt(i) == '#')
                count++;
            else
                break;
        }
        return count;
    }

    public static boolean isAsciiLike(char ch) {
        return ch >= 0x20 && ch <= 0x7E;
    }

    // "1. " / "12.\t" のような形式を番号付きリストとして判定（正規表現なし）
    // 条件：先頭が数字+、続いて '.'、続いて空白（スペース/タブ等）が1文字以上
    public static boolean isNumberedListLine(String trimmed) {
        if (trimmed == null || trimmed.isEmpty()) {
            return false;
        }

        int n = trimmed.length();
        int i = 0;

        char c0 = trimmed.charAt(0);
        if (c0 < '0' || c0 > '9') {
            return false;
        }

        // 数字列を読む
        while (i < n) {
            char ch = trimmed.charAt(i);
            if (ch < '0' || ch > '9') {
                break;
            }
            i++;
        }

        // 数字の後に '.' が必要
        if (i >= n || trimmed.charAt(i) != '.') {
            return false;
        }
        i++;

        // '.' の後に空白が1文字以上必要（元の \s+ と同じ）
        if (i >= n || !Character.isWhitespace(trimmed.charAt(i))) {
            return false;
        }

        return true; // 末尾まで空白でも OK（元の ".*" は空でもマッチする）
    }

    public static String replaceBrOutsideInlineCode(String s, String replacement) {
        if (s == null || s.isEmpty())
            return s;

        StringBuilder out = new StringBuilder(s.length());
        boolean inCode = false;

        for (int i = 0; i < s.length();) {
            char ch = s.charAt(i);

            // インラインコードは `<br>` を触らない
            if (ch == '`') {
                inCode = !inCode;
                out.append(ch);
                i++;
                continue;
            }

            if (!inCode) {
                int brLen = matchBrTagLen(s, i);
                if (brLen > 0) {
                    out.append(replacement);
                    i += brLen;
                    continue;
                }
            }

            out.append(ch);
            i++;
        }
        return out.toString();
    }

    public static int matchBrTagLen(String s, int i) {
        int n = s.length();
        if (i < 0 || i + 3 >= n)
            return 0;
        if (s.charAt(i) != '<')
            return 0;

        // "<br" / "<BR" / "<bR" / "<Br" を許可
        char b = s.charAt(i + 1);
        char r = s.charAt(i + 2);
        if (Character.toLowerCase(b) != 'b' || Character.toLowerCase(r) != 'r')
            return 0;

        int j = i + 3;

        // 任意の空白
        while (j < n && Character.isWhitespace(s.charAt(j)))
            j++;

        // 任意の "/"（<br/> or <br />）
        if (j < n && s.charAt(j) == '/') {
            j++;
            while (j < n && Character.isWhitespace(s.charAt(j)))
                j++;
        }

        // ">" で閉じる
        if (j < n && s.charAt(j) == '>') {
            return (j - i) + 1;
        }
        return 0;
    }

    // 連結用：空白を 1 個に寄せたい時だけ使う（テーブル用途）
    public static String collapseSpaces(String s) {
        if (s == null)
            return null;
        StringBuilder out = new StringBuilder(s.length());
        boolean prevSpace = false;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            boolean sp = Character.isWhitespace(ch);
            if (sp) {
                if (!prevSpace)
                    out.append(' ');
                prevSpace = true;
            } else {
                out.append(ch);
                prevSpace = false;
            }
        }
        // trim
        int len = out.length();
        int st = 0;
        while (st < len && out.charAt(st) == ' ')
            st++;
        int ed = len;
        while (ed > st && out.charAt(ed - 1) == ' ')
            ed--;
        return out.substring(st, ed);
    }
}
