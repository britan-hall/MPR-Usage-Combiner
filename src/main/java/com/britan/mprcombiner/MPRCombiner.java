package com.britan.mprcombiner;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public final class MPRCombiner {
    private static final String COL_SOURCE_FILE = "__source_file";
    private static final String COL_SOURCE_SHEET = "__source_sheet";
    private static final String COL_SOURCE_ROW = "__source_row";

    private static final Set<String> EXCEL_EXTENSIONS = Set.of("xlsx", "xlsm", "xls");

    /**
     * Maps a "normalized key" (letters+digits only, lowercased) to a canonical header name.
     * This lets us recover from headers that have extra whitespace, line breaks, punctuation,
     * or minor typos (e.g. "rport mo. nth" -> "Report Month").
     */
    private static final Map<String, String> CANONICAL_HEADERS = buildCanonicalHeaderMap();

    private MPRCombiner() {
    }

    public static void main(String[] args) throws Exception {
        Options options = buildOptions();
        CommandLine cli = parseArgs(options, args);

        if (cli.hasOption("help")) {
            printHelp(options);
            return;
        }

        Path inputDir = Paths.get(cli.getOptionValue("input")).toAbsolutePath().normalize();
        Path outputFile = Paths.get(cli.getOptionValue("output")).toAbsolutePath().normalize();
        String sheetName = cli.getOptionValue("sheet", "Usage");
        int headerRow1Based = Integer.parseInt(cli.getOptionValue("headerRow", "1"));
        boolean recursive = Boolean.parseBoolean(cli.getOptionValue("recursive", "false"));
        boolean diagnose = Boolean.parseBoolean(cli.getOptionValue("diagnose", "false"));
        boolean includeSource = Boolean.parseBoolean(cli.getOptionValue("includeSource", "false"));

        if (!Files.isDirectory(inputDir)) {
            throw new IllegalArgumentException("--input must be a directory: " + inputDir);
        }
        if (headerRow1Based < 1) {
            throw new IllegalArgumentException("--headerRow must be >= 1");
        }

        List<Path> excelFiles = listExcelFiles(inputDir, recursive);
        if (excelFiles.isEmpty()) {
            throw new IllegalArgumentException("No Excel files found in: " + inputDir);
        }

        List<UsageRow> allRows = new ArrayList<>();
        LinkedHashSet<String> unifiedColumns = new LinkedHashSet<>();

        DataFormatter formatter = new DataFormatter(Locale.US, true);

        int filesOpened = 0;
        int filesSkippedUnreadable = 0;
        int filesSkippedNoSheet = 0;
        int filesSkippedNoHeader = 0;
        int filesWithZeroRows = 0;

        for (Path excelPath : excelFiles) {
            try (InputStream in = Files.newInputStream(excelPath);
                 Workbook wb = WorkbookFactory.create(in)) {
                filesOpened++;
                Sheet sheet = findUsageSheetWithFallback(wb, sheetName);
                if (sheet == null) {
                    filesSkippedNoSheet++;
                    if (diagnose) {
                        System.err.println("[diagnose] No sheet match in " + excelPath.getFileName()
                                + " (looking for '" + sheetName + "'" + sheetFallbackNote(sheetName) + "). Sheets: " + listSheetNames(wb));
                    }
                    continue;
                }

                int headerRowIdx = headerRow1Based - 1;
                Row headerRow = sheet.getRow(headerRowIdx);
                if (headerRow == null) {
                    filesSkippedNoHeader++;
                    if (diagnose) {
                        System.err.println("[diagnose] Missing header row " + headerRow1Based + " in "
                                + excelPath.getFileName() + " sheet '" + sheet.getSheetName() + "'");
                    }
                    continue;
                }

                Map<Integer, String> colIndexToName = readHeader(headerRow, formatter);
                if (colIndexToName.isEmpty()) {
                    filesSkippedNoHeader++;
                    if (diagnose) {
                        System.err.println("[diagnose] Empty header on row " + headerRow1Based + " in "
                                + excelPath.getFileName() + " sheet '" + sheet.getSheetName() + "'");
                    }
                    continue;
                }
                unifiedColumns.addAll(colIndexToName.values());

                int lastRowNum = sheet.getLastRowNum();
                int extractedForFile = 0;
                for (int r = headerRowIdx + 1; r <= lastRowNum; r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) {
                        continue;
                    }
                    if (isRowEmpty(row)) {
                        continue;
                    }

                    Map<String, String> values = new LinkedHashMap<>();
                    for (Map.Entry<Integer, String> e : colIndexToName.entrySet()) {
                        int c = e.getKey();
                        String colName = e.getValue();
                        String cellValue = cellToString(row.getCell(c), formatter);
                        values.put(colName, cellValue);
                    }

                    allRows.add(new UsageRow(
                            excelPath.getFileName().toString(),
                            sheet.getSheetName(),
                            r + 1,
                            values
                    ));
                    extractedForFile++;
                }
                if (diagnose) {
                    System.err.println("[diagnose] Extracted " + extractedForFile + " rows from " + excelPath.getFileName()
                            + " sheet '" + sheet.getSheetName() + "' (header columns=" + colIndexToName.size() + ")");
                }
                if (extractedForFile == 0) {
                    filesWithZeroRows++;
                }
            } catch (Exception e) {
                System.err.println("Skipping unreadable file: " + excelPath + " (" + e.getMessage() + ")");
                filesSkippedUnreadable++;
            }
        }

        if (allRows.isEmpty()) {
            throw new IllegalStateException("No usage rows found. Confirm the sheet name and header row.");
        }

        List<String> outputColumns = new ArrayList<>();
        if (includeSource) {
            outputColumns.add(COL_SOURCE_FILE);
            outputColumns.add(COL_SOURCE_SHEET);
            outputColumns.add(COL_SOURCE_ROW);
        }
        outputColumns.addAll(unifiedColumns);

        Files.createDirectories(outputFile.getParent() == null ? Paths.get(".") : outputFile.getParent());
        if (isXlsxOutput(outputFile)) {
            writeXlsx(outputFile, outputColumns, allRows);
        } else {
            writeCsv(outputFile, outputColumns, allRows);
        }

        System.out.println("Combined " + allRows.size() + " rows from " + excelFiles.size() + " files into: " + outputFile);
        if (diagnose) {
            System.err.println("[diagnose] Files found: " + excelFiles.size()
                    + ", opened: " + filesOpened
                    + ", skipped (unreadable): " + filesSkippedUnreadable
                    + ", skipped (no sheet): " + filesSkippedNoSheet
                    + ", skipped (no header): " + filesSkippedNoHeader
                    + ", opened with 0 extracted rows: " + filesWithZeroRows);
        }
    }

    private static Options buildOptions() {
        Options options = new Options();
        options.addOption(Option.builder().longOpt("input").hasArg().argName("DIR").desc("Input directory of Excel files").build());
        options.addOption(Option.builder().longOpt("output").hasArg().argName("FILE").desc("Output CSV file").build());
        options.addOption(Option.builder().longOpt("sheet").hasArg().argName("NAME").desc("Sheet name to extract (default: Usage)").build());
        options.addOption(Option.builder().longOpt("headerRow").hasArg().argName("N").desc("1-based header row index (default: 1)").build());
        options.addOption(Option.builder().longOpt("recursive").hasArg().argName("true|false").desc("Search subfolders (default: false)").build());
        options.addOption(Option.builder().longOpt("includeSource").hasArg().argName("true|false")
                .desc("Include __source_* columns in combined output (default: false)").build());
        options.addOption(Option.builder().longOpt("diagnose").hasArg().argName("true|false")
                .desc("Print per-file extraction diagnostics to stderr (default: false)").build());
        options.addOption(Option.builder("h").longOpt("help").desc("Show help").build());
        return options;
    }

    private static CommandLine parseArgs(Options options, String[] args) throws ParseException {
        CommandLineParser parser = new DefaultParser();
        CommandLine cli = parser.parse(options, args);

        if (cli.hasOption("help")) {
            return cli;
        }

        List<String> missing = new ArrayList<>();
        if (!cli.hasOption("input")) missing.add("--input");
        if (!cli.hasOption("output")) missing.add("--output");
        if (!missing.isEmpty()) {
            throw new ParseException("Missing required option(s): " + String.join(", ", missing));
        }
        return cli;
    }

    private static void printHelp(Options options) {
        HelpFormatter hf = new HelpFormatter();
        hf.setWidth(110);
        hf.printHelp("mpr-combiner", options, true);
    }

    private static List<Path> listExcelFiles(Path inputDir, boolean recursive) throws IOException {
        int maxDepth = recursive ? Integer.MAX_VALUE : 1;
        try (Stream<Path> stream = Files.walk(inputDir, maxDepth)) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(p -> isExcelFile(p.getFileName().toString()))
                    .sorted(Comparator.comparing(Path::toString))
                    .collect(Collectors.toList());
        }
    }

    private static boolean isExcelFile(String name) {
        int dot = name.lastIndexOf('.');
        if (dot < 0) return false;
        String ext = name.substring(dot + 1).toLowerCase(Locale.ROOT);
        return EXCEL_EXTENSIONS.contains(ext);
    }

    private static Sheet findSheetByName(Workbook wb, String targetName) {
        String normTarget = normalizeSheetName(targetName);
        if (normTarget.isEmpty()) return null;

        Sheet exact = wb.getSheet(targetName);
        if (exact != null) return exact;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet s = wb.getSheetAt(i);
            if (s == null || s.getSheetName() == null) continue;
            String normSheet = normalizeSheetName(s.getSheetName());
            if (!normSheet.isEmpty() && normSheet.equalsIgnoreCase(normTarget)) {
                return s;
            }
        }
        return null;
    }

    private static Sheet findUsageSheetWithFallback(Workbook wb, String requestedSheetName) {
        Sheet primary = findSheetByName(wb, requestedSheetName);
        if (primary != null) return primary;

        // Option A: if the user is asking for the default Usage sheet and it doesn't exist,
        // fall back to "Usage (Part B)" which is common in other MPR templates.
        String normRequested = normalizeSheetName(requestedSheetName);
        if (normRequested.equalsIgnoreCase("Usage")) {
            return findSheetByName(wb, "Usage (Part B)");
        }
        return null;
    }

    private static String sheetFallbackNote(String requestedSheetName) {
        String normRequested = normalizeSheetName(requestedSheetName);
        if (normRequested.equalsIgnoreCase("Usage")) {
            return " or 'Usage (Part B)'";
        }
        return "";
    }

    private static String normalizeSheetName(String s) {
        if (s == null) return "";
        String t = s.trim();
        if (t.isEmpty()) return "";
        // Collapse all internal whitespace to a single space
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

    private static String listSheetNames(Workbook wb) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet s = wb.getSheetAt(i);
            if (s != null && s.getSheetName() != null) {
                names.add("'" + s.getSheetName() + "'");
            }
        }
        return String.join(", ", names);
    }

    private static Map<Integer, String> readHeader(Row headerRow, DataFormatter formatter) {
        Map<Integer, String> map = new LinkedHashMap<>();
        short lastCellNum = headerRow.getLastCellNum();
        if (lastCellNum < 0) return map;

        for (int c = 0; c < lastCellNum; c++) {
            String raw = cellToString(headerRow.getCell(c), formatter);
            String name = normalizeHeaderName(raw);
            if (name == null) {
                continue;
            }

            String unique = name;
            int suffix = 2;
            while (map.containsValue(unique)) {
                unique = name + "_" + suffix;
                suffix++;
            }
            map.put(c, unique);
        }
        return map;
    }

    private static String normalizeHeaderName(String raw) {
        if (raw == null) return null;

        String s = raw.replace("\u00A0", " "); // non-breaking space
        s = s.trim();
        if (s.isEmpty()) return null;

        String key = toHeaderKey(s);
        if (!key.isEmpty()) {
            String canonical = CANONICAL_HEADERS.get(key);
            if (canonical != null) return canonical;

            // A tiny bit of typo forgiveness for known troublesome headers.
            // Example: "rportmonth" -> "reportmonth"
            if ("rportmonth".equals(key)) return "Report Month";
            if ("reportunitmonth".equals(key)) return "Report Month";
        }

        // Fall back to a cleaned version of whatever we got.
        return collapseWhitespace(s);
    }

    private static Map<String, String> buildCanonicalHeaderMap() {
        Map<String, String> m = new LinkedHashMap<>();

        // Upload columns
        putCanonical(m, "Processor ID", "processorid", "processor id", "processoridnumber");
        putCanonical(m, "Processor Name", "processorname", "processor name");
        putCanonical(m, "Report Month", "reportmonth", "report month", "reportmo", "report mo", "month");
        putCanonical(m, "Report Year", "reportyear", "report year", "year");
        putCanonical(m, "State", "state", "st");
        putCanonical(m, "Recipient Agency Number", "recipientagencynumber", "recipient agency number", "ra number", "ranumber", "recipient agency #", "recipientagency#");
        putCanonical(m, "Recipient Agency Name", "recipientagencyname", "recipient agency name", "ra name", "raname");
        putCanonical(m, "Product Number", "productnumber", "product number", "productnbr", "product nbr", "product #", "product#");
        putCanonical(m, "Product Name", "productname", "product name");
        putCanonical(m, "USDA Material", "usdamaterial", "usda material", "usda matl", "usda");
        putCanonical(m, "EPDS DF", "epdsdf", "epds df", "df");
        putCanonical(m, "Case Qty", "caseqty", "case qty", "cases", "case quantity");
        putCanonical(m, "Used LBS", "usedlbs", "used lbs", "usedlbs.", "used pounds", "lbs used");

        // Common extra columns
        putCanonical(m, "COOP", "coop");

        return m;
    }

    private static void putCanonical(Map<String, String> m, String canonical, String... variants) {
        m.put(toHeaderKey(canonical), canonical);
        for (String v : variants) {
            m.put(toHeaderKey(v), canonical);
        }
    }

    private static String toHeaderKey(String s) {
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

    private static String collapseWhitespace(String s) {
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

    private static boolean isRowEmpty(Row row) {
        short last = row.getLastCellNum();
        if (last < 0) return true;
        for (int c = 0; c < last; c++) {
            Cell cell = row.getCell(c);
            if (cell == null) continue;
            CellType type = cell.getCellType();
            if (type == CellType.BLANK) continue;
            if (type == CellType.STRING && cell.getStringCellValue() != null && !cell.getStringCellValue().trim().isEmpty()) {
                return false;
            }
            if (type != CellType.STRING && type != CellType.BLANK) {
                String v = cellToString(cell, new DataFormatter(Locale.US, true));
                if (v != null && !v.trim().isEmpty()) return false;
            }
        }
        return true;
    }

    private static String cellToString(Cell cell, DataFormatter formatter) {
        if (cell == null) return "";
        try {
            return Objects.toString(formatter.formatCellValue(cell), "");
        } catch (Exception e) {
            return "";
        }
    }

    private static boolean isXlsxOutput(Path outputFile) {
        String name = outputFile.getFileName() == null ? "" : outputFile.getFileName().toString();
        return name.toLowerCase(Locale.ROOT).endsWith(".xlsx");
    }

    private static void writeCsv(Path outputFile, List<String> columns, List<UsageRow> rows) throws IOException {
        try (BufferedWriter w = Files.newBufferedWriter(outputFile, StandardCharsets.UTF_8)) {
            w.write(columns.stream().map(MPRCombiner::csvEscape).collect(Collectors.joining(",")));
            w.write("\n");

            for (UsageRow r : rows) {
                List<String> vals = new ArrayList<>(columns.size());
                if (columns.size() >= 3
                        && COL_SOURCE_FILE.equals(columns.get(0))
                        && COL_SOURCE_SHEET.equals(columns.get(1))
                        && COL_SOURCE_ROW.equals(columns.get(2))) {
                    vals.add(r.sourceFile);
                    vals.add(r.sourceSheet);
                    vals.add(Integer.toString(r.sourceRow));
                    for (int i = 3; i < columns.size(); i++) {
                        String col = columns.get(i);
                        vals.add(r.valuesByColumn.getOrDefault(col, ""));
                    }
                } else {
                    for (String col : columns) {
                        vals.add(r.valuesByColumn.getOrDefault(col, ""));
                    }
                }
                w.write(vals.stream().map(MPRCombiner::csvEscape).collect(Collectors.joining(",")));
                w.write("\n");
            }
        }
    }

    private static void writeXlsx(Path outputFile, List<String> columns, List<UsageRow> rows) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Usage");

            Row header = sheet.createRow(0);
            for (int c = 0; c < columns.size(); c++) {
                header.createCell(c).setCellValue(columns.get(c));
            }

            for (int rIdx = 0; rIdx < rows.size(); rIdx++) {
                UsageRow r = rows.get(rIdx);
                Row row = sheet.createRow(rIdx + 1);

                if (columns.size() >= 3
                        && COL_SOURCE_FILE.equals(columns.get(0))
                        && COL_SOURCE_SHEET.equals(columns.get(1))
                        && COL_SOURCE_ROW.equals(columns.get(2))) {
                    row.createCell(0).setCellValue(r.sourceFile);
                    row.createCell(1).setCellValue(r.sourceSheet);
                    row.createCell(2).setCellValue(r.sourceRow);
                    for (int c = 3; c < columns.size(); c++) {
                        row.createCell(c).setCellValue(r.valuesByColumn.getOrDefault(columns.get(c), ""));
                    }
                } else {
                    for (int c = 0; c < columns.size(); c++) {
                        row.createCell(c).setCellValue(r.valuesByColumn.getOrDefault(columns.get(c), ""));
                    }
                }
            }

            // Keep file size reasonable; autosizing thousands of columns/rows can be slow.
            try (OutputStream os = Files.newOutputStream(outputFile)) {
                wb.write(os);
            }
        }
    }

    private static String csvEscape(String s) {
        if (s == null) return "";
        boolean needsQuotes = s.contains(",") || s.contains("\"") || s.contains("\n") || s.contains("\r");
        if (!needsQuotes) return s;
        return "\"" + s.replace("\"", "\"\"") + "\"";
    }

    private static final class UsageRow {
        final String sourceFile;
        final String sourceSheet;
        final int sourceRow;
        final Map<String, String> valuesByColumn;

        UsageRow(String sourceFile, String sourceSheet, int sourceRow, Map<String, String> valuesByColumn) {
            this.sourceFile = sourceFile;
            this.sourceSheet = sourceSheet;
            this.sourceRow = sourceRow;
            this.valuesByColumn = valuesByColumn;
        }
    }
}

