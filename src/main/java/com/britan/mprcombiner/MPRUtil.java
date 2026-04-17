package com.britan.mprcombiner;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

final class MPRUtil {
    private MPRUtil() {
    }

    static boolean isXlsxOutput(Path outputFile) {
        String name = outputFile.getFileName() == null ? "" : outputFile.getFileName().toString();
        return name.toLowerCase(Locale.ROOT).endsWith(".xlsx");
    }

    static String normalizeSheetName(String s) {
        if (s == null) return "";
        String t = s.trim();
        if (t.isEmpty()) return "";
        return collapseWhitespace(t);
    }

    static String listSheetNames(Workbook wb) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet s = wb.getSheetAt(i);
            if (s != null && s.getSheetName() != null) {
                names.add("'" + s.getSheetName() + "'");
            }
        }
        return String.join(", ", names);
    }

    static String toHeaderKey(String s) {
        if (s == null) return "";
        String t = collapseWhitespace(s).toLowerCase(Locale.ROOT);
        StringBuilder b = new StringBuilder(t.length());
        for (int i = 0; i < t.length(); i++) {
            char ch = t.charAt(i);
            if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) {
                b.append(ch);
            }
        }
        return b.toString();
    }

    static String collapseWhitespace(String s) {
        if (s == null) return "";
        String t = s.trim();
        StringBuilder b = new StringBuilder(t.length());
        boolean lastWasSpace = false;
        for (int i = 0; i < t.length(); i++) {
            char ch = t.charAt(i);
            boolean isSpace = Character.isWhitespace(ch);
            if (isSpace) {
                if (!lastWasSpace) {
                    b.append(' ');
                    lastWasSpace = true;
                }
            } else {
                b.append(ch);
                lastWasSpace = false;
            }
        }
        return b.toString();
    }

    static String csvEscape(String s) {
        if (s == null) return "";
        boolean needsQuotes = s.contains(",") || s.contains("\"") || s.contains("\n") || s.contains("\r");
        if (!needsQuotes) return s;
        return "\"" + s.replace("\"", "\"\"") + "\"";
    }
}

