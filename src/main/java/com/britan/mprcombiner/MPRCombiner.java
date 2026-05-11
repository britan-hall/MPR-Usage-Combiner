package com.britan.mprcombiner;

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
import java.time.LocalDate;
import java.time.Month;
import java.time.temporal.IsoFields;
import java.time.Year;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.OptionalInt;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public final class MPRCombiner {

    /**
     * @param outputFile when {@code null}, writes {@code 99_combinedUsage-<year>-Q<n>-<state>-MPR.xlsx} using the most common
     *                    Report Year and State from the combined rows. Calendar quarter comes from the first {@code Q1}–{@code Q4}
     *                    found in the combine folder path (leaf toward root); if none is found, quarter is inferred from
     *                    Report Month. The file goes in {@code outputDirectoryWhenAuto} when that is non-null, otherwise in
     *                    {@code inputDir}.
     */
    public record CombineParams(
            Path inputDir,
            Path outputFile,
            Path outputDirectoryWhenAuto,
            String sheetName,
            int headerRow1Based,
            boolean recursive,
            boolean diagnose,
            boolean includeSource
    ) {
    }

    private static final String COL_SOURCE_FILE = "__source_file";
    private static final String COL_SOURCE_SHEET = "__source_sheet";
    private static final String COL_SOURCE_ROW = "__source_row";

    private static final Set<String> EXCEL_EXTENSIONS = Set.of("xlsx", "xlsm", "xls");

    /** Max data rows per Usage tab when combined output exceeds this count (header row is extra). */
    static final int MAX_USAGE_DATA_ROWS_PER_SHEET = 25_000;

    /**
     * Maps a "normalized key" (letters+digits only, lowercased) to a canonical header name.
     * This lets us recover from headers that have extra whitespace, line breaks, punctuation,
     * or minor typos (e.g. "rport mo. nth" -> "Report Month").
     */
    private static final Map<String, String> CANONICAL_HEADERS = buildCanonicalHeaderMap();

    private MPRCombiner() {
    }

    public static void main(String[] args) {
        MPRCombinerGui.launch();
    }

    /**
     * Combines Usage sheets from Excel files under {@code inputDir}.
     *
     * @param outputFile {@code null} for auto-named {@code 99_combinedUsage-…MPR.xlsx} (see {@code outputDirectoryWhenAuto})
     * @param outputDirectoryWhenAuto when {@code outputFile} is {@code null} and this is non-null, the auto-named file
     *                                is written here; when {@code null}, it is written in {@code inputDir}
     * @return path of the written file
     */
    public static Path runLocalCombine(
            Path inputDir,
            Path outputFile,
            Path outputDirectoryWhenAuto,
            String sheetName,
            int headerRow1Based,
            boolean recursive,
            boolean diagnose,
            boolean includeSource
    ) throws Exception {
        return combine(new CombineParams(
                inputDir, outputFile, outputDirectoryWhenAuto, sheetName, headerRow1Based, recursive, diagnose, includeSource));
    }

    static Path combine(CombineParams params) throws Exception {
        Path inputDir = params.inputDir().toAbsolutePath().normalize();
        String sheetName = params.sheetName();
        int headerRow1Based = params.headerRow1Based();
        boolean recursive = params.recursive();
        boolean diagnose = params.diagnose();
        boolean includeSource = params.includeSource();

        if (!Files.isDirectory(inputDir)) {
            throw new IllegalArgumentException("Input must be a directory: " + inputDir);
        }
        if (headerRow1Based < 1) {
            throw new IllegalArgumentException("Header row must be >= 1");
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
                                + " (looking for '" + sheetName + "'" + sheetFallbackNote(sheetName) + "). Sheets: " + MPRUtil.listSheetNames(wb));
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

        Path resolvedOutput = params.outputFile();
        if (resolvedOutput == null) {
            String stem = buildAutoOutputStem(inputDir, allRows);
            Path autoDir = params.outputDirectoryWhenAuto();
            if (autoDir != null) {
                autoDir = autoDir.toAbsolutePath().normalize();
                if (!Files.isDirectory(autoDir)) {
                    throw new IllegalArgumentException("Output folder must be an existing directory: " + autoDir);
                }
            } else {
                autoDir = inputDir;
            }
            resolvedOutput = uniquifyInDirectory(autoDir, "99_combined" + stem + ".xlsx");
        } else {
            resolvedOutput = resolvedOutput.toAbsolutePath().normalize();
        }

        List<String> outputColumns = new ArrayList<>();
        if (includeSource) {
            outputColumns.add(COL_SOURCE_FILE);
            outputColumns.add(COL_SOURCE_SHEET);
            outputColumns.add(COL_SOURCE_ROW);
        }
        outputColumns.addAll(unifiedColumns);

        Files.createDirectories(resolvedOutput.getParent() == null ? Paths.get(".") : resolvedOutput.getParent());
        if (MPRUtil.isXlsxOutput(resolvedOutput)) {
            writeXlsx(resolvedOutput, outputColumns, allRows);
        } else {
            writeCsv(resolvedOutput, outputColumns, allRows);
        }

        System.out.println("Combined " + allRows.size() + " rows from " + excelFiles.size() + " files into: " + resolvedOutput);
        if (diagnose) {
            System.err.println("[diagnose] Files found: " + excelFiles.size()
                    + ", opened: " + filesOpened
                    + ", skipped (unreadable): " + filesSkippedUnreadable
                    + ", skipped (no sheet): " + filesSkippedNoSheet
                    + ", skipped (no header): " + filesSkippedNoHeader
                    + ", opened with 0 extracted rows: " + filesWithZeroRows);
        }
        return resolvedOutput;
    }

    private static final Pattern QUARTER_IN_SEGMENT = Pattern.compile("(?i)Q([1-4])(?!\\d)");

    /**
     * Middle part of auto filename (between {@code 99_combined} and {@code .xlsx}):
     * {@code Usage-<year>-Q<quarter>-<state>-MPR}. Quarter prefers the first {@code Q1}–{@code Q4} in the combine folder path;
     * otherwise it follows Report Month in the rows.
     */
    static String buildAutoOutputStem(Path inputDir, List<UsageRow> rows) {
        String year = dominantToken(rows, "Report Year", () -> String.valueOf(Year.now().getValue()));
        year = normalizeYearForFilename(year);

        String state = dominantToken(rows, "State", () -> "UNK");
        state = sanitizeFileToken(state, 12);

        int quarter = quarterFromFolderPath(inputDir).orElseGet(() -> dominantCalendarQuarter(rows));
        return "Usage-" + year + "-Q" + quarter + "-" + state + "-MPR";
    }

    /**
     * Walk from the combine folder toward the filesystem root; first segment containing {@code Q1}–{@code Q4} wins
     * ({@code Q} must not be followed by another digit, so {@code Q12} does not match as quarter 1).
     */
    private static OptionalInt quarterFromFolderPath(Path inputDir) {
        Path walk = inputDir.toAbsolutePath().normalize();
        while (walk != null) {
            Path name = walk.getFileName();
            if (name != null) {
                OptionalInt q = quarterFromPathSegment(name.toString());
                if (q.isPresent()) {
                    return q;
                }
            }
            walk = walk.getParent();
        }
        return OptionalInt.empty();
    }

    static OptionalInt quarterFromPathSegment(String segment) {
        if (segment == null || segment.isEmpty()) {
            return OptionalInt.empty();
        }
        Matcher m = QUARTER_IN_SEGMENT.matcher(segment);
        if (m.find()) {
            return OptionalInt.of(Integer.parseInt(m.group(1)));
        }
        return OptionalInt.empty();
    }

    private static String dominantToken(List<UsageRow> rows, String column, java.util.function.Supplier<String> fallback) {
        Map<String, Integer> counts = new HashMap<>();
        for (UsageRow r : rows) {
            String v = r.valuesByColumn.getOrDefault(column, "").trim();
            if (!v.isEmpty()) {
                counts.merge(v, 1, Integer::sum);
            }
        }
        if (counts.isEmpty()) {
            return sanitizeFileToken(fallback.get(), column.length() > 8 ? 12 : 8);
        }
        String best = counts.entrySet().stream()
                .max(Map.Entry.<String, Integer>comparingByValue()
                        .thenComparing(e -> e.getKey().length())
                        .thenComparing(Map.Entry::getKey))
                .map(Map.Entry::getKey)
                .orElse(fallback.get());
        return best.trim();
    }

    private static String normalizeYearForFilename(String raw) {
        String s = raw.trim();
        if (s.matches("\\d{4}")) {
            return s;
        }
        if (s.contains(".")) {
            try {
                int y = (int) Math.round(Double.parseDouble(s));
                if (y >= 1900 && y <= 2200) {
                    return String.valueOf(y);
                }
            } catch (NumberFormatException ignored) {
            }
        }
        Matcher m = Pattern.compile("(19|20)\\d{2}").matcher(s);
        if (m.find()) {
            return m.group();
        }
        return sanitizeFileToken(String.valueOf(Year.now().getValue()), 4);
    }

    private static int dominantCalendarQuarter(List<UsageRow> rows) {
        int[] counts = new int[5];
        for (UsageRow r : rows) {
            parseReportMonth(r.valuesByColumn.get("Report Month")).ifPresent(month -> {
                int q = (month - 1) / 3 + 1;
                if (q >= 1 && q <= 4) {
                    counts[q]++;
                }
            });
        }
        int fallback = LocalDate.now().get(IsoFields.QUARTER_OF_YEAR);
        int bestQ = fallback;
        int bestCount = 0;
        for (int q = 1; q <= 4; q++) {
            if (counts[q] > bestCount) {
                bestCount = counts[q];
                bestQ = q;
            }
        }
        return bestQ;
    }

    private static OptionalInt parseReportMonth(String raw) {
        if (raw == null) {
            return OptionalInt.empty();
        }
        String s = raw.trim();
        if (s.isEmpty()) {
            return OptionalInt.empty();
        }
        if (s.matches("-?\\d+(\\.\\d+)?")) {
            try {
                int n = (int) Math.round(Double.parseDouble(s));
                if (n >= 1 && n <= 12) {
                    return OptionalInt.of(n);
                }
            } catch (NumberFormatException ignored) {
            }
        }
        String key = s.toLowerCase(Locale.ROOT).replace(".", "");
        for (Month mo : Month.values()) {
            String full = mo.name().toLowerCase(Locale.ROOT);
            String short3 = full.length() >= 3 ? full.substring(0, 3) : full;
            if (key.equals(full) || key.startsWith(short3)) {
                return OptionalInt.of(mo.getValue());
            }
        }
        return OptionalInt.empty();
    }

    private static String sanitizeFileToken(String raw, int maxLen) {
        String s = raw.trim().toUpperCase(Locale.ROOT).replaceAll("[^A-Z0-9_-]+", "_");
        if (s.isEmpty()) {
            return "UNK";
        }
        if (s.length() > maxLen) {
            return s.substring(0, maxLen);
        }
        return s;
    }

    private static Path uniquifyInDirectory(Path dir, String filename) {
        Path candidate = dir.resolve(filename).normalize();
        if (!Files.exists(candidate)) {
            return candidate;
        }
        String base = filename;
        String ext = "";
        int dot = filename.lastIndexOf('.');
        if (dot > 0) {
            base = filename.substring(0, dot);
            ext = filename.substring(dot);
        }
        for (int i = 2; i < 10_000; i++) {
            Path p = dir.resolve(base + "_" + i + ext).normalize();
            if (!Files.exists(p)) {
                return p;
            }
        }
        throw new IllegalStateException("Could not find a free filename in: " + dir);
    }

    private static List<Path> listExcelFiles(Path inputDir, boolean recursive) throws IOException {
        int maxDepth = recursive ? Integer.MAX_VALUE : 1;
        try (Stream<Path> stream = Files.walk(inputDir, maxDepth)) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(p -> isExcelFile(p.getFileName().toString()))
                    .filter(p -> !isAutoCombinedOutputName(p.getFileName().toString()))
                    .sorted(Comparator.comparing(Path::toString))
                    .collect(Collectors.toList());
        }
    }

    /** Skips prior combined outputs so re-running in the same folder does not merge them as inputs. */
    private static boolean isAutoCombinedOutputName(String fileName) {
        String lower = fileName.toLowerCase(Locale.ROOT);
        return lower.startsWith("99_combined") && lower.endsWith(".xlsx");
    }

    private static boolean isExcelFile(String name) {
        int dot = name.lastIndexOf('.');
        if (dot < 0) return false;
        String ext = name.substring(dot + 1).toLowerCase(Locale.ROOT);
        return EXCEL_EXTENSIONS.contains(ext);
    }

    private static Sheet findSheetByName(Workbook wb, String targetName) {
        String normTarget = MPRUtil.normalizeSheetName(targetName);
        if (normTarget.isEmpty()) return null;

        Sheet exact = wb.getSheet(targetName);
        if (exact != null) return exact;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet s = wb.getSheetAt(i);
            if (s == null || s.getSheetName() == null) continue;
            String normSheet = MPRUtil.normalizeSheetName(s.getSheetName());
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
        String normRequested = MPRUtil.normalizeSheetName(requestedSheetName);
        if (normRequested.equalsIgnoreCase("Usage")) {
            return findSheetByName(wb, "Usage (Part B)");
        }
        return null;
    }

    private static String sheetFallbackNote(String requestedSheetName) {
        String normRequested = MPRUtil.normalizeSheetName(requestedSheetName);
        if (normRequested.equalsIgnoreCase("Usage")) {
            return " or 'Usage (Part B)'";
        }
        return "";
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

        String key = MPRUtil.toHeaderKey(s);
        if (!key.isEmpty()) {
            String canonical = CANONICAL_HEADERS.get(key);
            if (canonical != null) return canonical;

            // A tiny bit of typo forgiveness for known troublesome headers.
            // Example: "rportmonth" -> "reportmonth"
            if ("rportmonth".equals(key)) return "Report Month";
            if ("reportunitmonth".equals(key)) return "Report Month";
        }

        // Fall back to a cleaned version of whatever we got.
        return MPRUtil.collapseWhitespace(s);
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
        m.put(MPRUtil.toHeaderKey(canonical), canonical);
        for (String v : variants) {
            m.put(MPRUtil.toHeaderKey(v), canonical);
        }
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

    private static void writeCsv(Path outputFile, List<String> columns, List<UsageRow> rows) throws IOException {
        try (BufferedWriter w = Files.newBufferedWriter(outputFile, StandardCharsets.UTF_8)) {
            w.write(columns.stream().map(MPRUtil::csvEscape).collect(Collectors.joining(",")));
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
                w.write(vals.stream().map(MPRUtil::csvEscape).collect(Collectors.joining(",")));
                w.write("\n");
            }
        }
    }

    private static void writeXlsx(Path outputFile, List<String> columns, List<UsageRow> rows) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            int chunkCount = usageSheetChunkCount(rows.size());
            for (int chunk = 0; chunk < chunkCount; chunk++) {
                int from = chunk * MAX_USAGE_DATA_ROWS_PER_SHEET;
                int to = Math.min(from + MAX_USAGE_DATA_ROWS_PER_SHEET, rows.size());
                Sheet sheet = wb.createSheet(usageSheetName(chunk, chunkCount));
                writeUsageSheetHeader(sheet, columns);
                writeUsageSheetRows(sheet, columns, rows, from, to);
            }

            // Keep file size reasonable; autosizing thousands of columns/rows can be slow.
            try (OutputStream os = Files.newOutputStream(outputFile)) {
                wb.write(os);
            }
        }
    }

    static int usageSheetChunkCount(int dataRowCount) {
        if (dataRowCount <= 0) {
            return 0;
        }
        if (dataRowCount <= MAX_USAGE_DATA_ROWS_PER_SHEET) {
            return 1;
        }
        return (dataRowCount + MAX_USAGE_DATA_ROWS_PER_SHEET - 1) / MAX_USAGE_DATA_ROWS_PER_SHEET;
    }

    static String usageSheetName(int chunkIndexZeroBased, int chunkCount) {
        if (chunkCount <= 1) {
            return "Usage";
        }
        return "Usage" + (chunkIndexZeroBased + 1);
    }

    private static void writeUsageSheetHeader(Sheet sheet, List<String> columns) {
        Row header = sheet.createRow(0);
        for (int c = 0; c < columns.size(); c++) {
            header.createCell(c).setCellValue(columns.get(c));
        }
    }

    private static void writeUsageSheetRows(
            Sheet sheet,
            List<String> columns,
            List<UsageRow> rows,
            int fromInclusive,
            int toExclusive
    ) {
        boolean withSource = columns.size() >= 3
                && COL_SOURCE_FILE.equals(columns.get(0))
                && COL_SOURCE_SHEET.equals(columns.get(1))
                && COL_SOURCE_ROW.equals(columns.get(2));

        for (int rIdx = fromInclusive; rIdx < toExclusive; rIdx++) {
            UsageRow r = rows.get(rIdx);
            Row row = sheet.createRow(rIdx - fromInclusive + 1);

            if (withSource) {
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

