package com.britan.mprcombiner;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertTrue;

public class MPRCombinerTest {

    @TempDir
    Path tempDir;

    @Test
    void combinesUsageAcrossFilesWithUnionOfHeaders() throws Exception {
        Path inDir = tempDir.resolve("input");
        Path outDir = tempDir.resolve("output");
        Files.createDirectories(inDir);
        Files.createDirectories(outDir);

        writeXlsx(inDir.resolve("a.xlsx"), new String[]{"Date", "kWh"}, new String[][]{
                {"2026-01-01", "10"},
                {"2026-01-02", "20"}
        });
        writeXlsx(inDir.resolve("b.xlsx"), new String[]{"Date", "Therms"}, new String[][]{
                {"2026-01-01", "5"}
        });

        Path out = outDir.resolve("combined.csv");
        MPRCombiner.main(new String[]{
                "--input", inDir.toString(),
                "--output", out.toString()
        });

        List<String> lines = readAllLinesUtf8(out);
        String header = lines.get(0);

        assertTrue(header.contains("Date"));
        assertTrue(header.contains("kWh"));
        assertTrue(header.contains("Therms"));

        String csv = String.join("\n", lines);
        assertTrue(csv.contains("2026-01-01"));
        assertTrue(csv.contains("2026-01-02"));
    }

    @Test
    void skipsWorkbookWithoutUsageSheet() throws Exception {
        Path inDir = tempDir.resolve("input2");
        Path out = tempDir.resolve("combined2.csv");
        Files.createDirectories(inDir);

        // workbook with a different sheet name
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("NotUsage");
            try (OutputStream os = Files.newOutputStream(inDir.resolve("no-usage.xlsx"))) {
                wb.write(os);
            }
        }

        writeXlsx(inDir.resolve("has-usage.xlsx"), new String[]{"A"}, new String[][]{{"1"}});

        MPRCombiner.main(new String[]{
                "--input", inDir.toString(),
                "--output", out.toString()
        });

        List<String> lines = readAllLinesUtf8(out);
        String csv = String.join("\n", lines);
        assertTrue(csv.contains("\n1\n") || csv.endsWith("\n1") || csv.contains(",1"));
    }

    @Test
    void fallsBackToUsagePartBWhenUsageNotPresent() throws Exception {
        Path inDir = tempDir.resolve("input_fallback");
        Path out = tempDir.resolve("fallback.csv");
        Files.createDirectories(inDir);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Usage (Part B)");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("Processor ID");
            var r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("P1");
            try (OutputStream os = Files.newOutputStream(inDir.resolve("partb.xlsx"))) {
                wb.write(os);
            }
        }

        MPRCombiner.main(new String[]{
                "--input", inDir.toString(),
                "--output", out.toString()
        });

        List<String> lines = readAllLinesUtf8(out);
        String csv = String.join("\n", lines);
        assertTrue(csv.contains("P1"));
    }

    @Test
    void normalizesBrokenHeadersLikeReportMonth() throws Exception {
        Path inDir = tempDir.resolve("input_headers");
        Path out = tempDir.resolve("headers.csv");
        Files.createDirectories(inDir);

        // Deliberately broken spacing/punctuation/typo in the header name.
        writeXlsx(inDir.resolve("broken.xlsx"),
                new String[]{"Processor ID", "Rport Mo.     nth", "Report Year"},
                new String[][]{
                        {"P1", "10", "2025"}
                });

        MPRCombiner.main(new String[]{
                "--input", inDir.toString(),
                "--output", out.toString()
        });

        List<String> lines = readAllLinesUtf8(out);
        String header = lines.get(0);
        assertTrue(header.contains("Report Month"));
        String csv = String.join("\n", lines);
        assertTrue(csv.contains(",10,")); // month value made it through
    }

    static Stream<Arguments> headerVariants() {
        return Stream.of(
                // Processor ID
                Arguments.of("Processor ID", "Processor ID", "P1"),
                Arguments.of("processor id", "Processor ID", "P1"),
                Arguments.of("Processor   ID", "Processor ID", "P1"),
                Arguments.of("Processor\nID", "Processor ID", "P1"),

                // Processor Name
                Arguments.of("Processor Name", "Processor Name", "Acme Foods"),
                Arguments.of("processor name", "Processor Name", "Acme Foods"),
                Arguments.of("Processor   Name", "Processor Name", "Acme Foods"),

                // Report Month
                Arguments.of("Report Month", "Report Month", "10"),
                Arguments.of("Report Mo", "Report Month", "10"),
                Arguments.of("Report Mo\tnth", "Report Month", "10"),
                Arguments.of("Report\nMonth", "Report Month", "10"),
                Arguments.of("Rport Mo.     nth", "Report Month", "10"),
                Arguments.of("Month", "Report Month", "10"),

                // Report Year
                Arguments.of("Report Year", "Report Year", "2025"),
                Arguments.of("Year", "Report Year", "2025"),
                Arguments.of("Report\nYear", "Report Year", "2025"),

                // State
                Arguments.of("State", "State", "AL"),
                Arguments.of("ST", "State", "AL"),

                // Recipient Agency Number
                Arguments.of("Recipient Agency Number", "Recipient Agency Number", "AL-158"),
                Arguments.of("RA Number", "Recipient Agency Number", "AL-158"),
                Arguments.of("Recipient Agency #", "Recipient Agency Number", "AL-158"),

                // Recipient Agency Name
                Arguments.of("Recipient Agency Name", "Recipient Agency Name", "Hoover City"),
                Arguments.of("RA Name", "Recipient Agency Name", "Hoover City"),

                // Product Number
                Arguments.of("Product Number", "Product Number", "1000002789"),
                Arguments.of("Product Nbr", "Product Number", "1000002789"),
                Arguments.of("Product #", "Product Number", "1000002789"),

                // Product Name
                Arguments.of("Product Name", "Product Name", "Tater Tots"),
                Arguments.of("Product\nName", "Product Name", "Tater Tots"),

                // USDA Material
                Arguments.of("USDA Material", "USDA Material", "100506"),
                Arguments.of("USDA", "USDA Material", "100506"),

                // EPDS DF
                Arguments.of("EPDS DF", "EPDS DF", "54.55"),
                Arguments.of("DF", "EPDS DF", "54.55"),

                // Case Qty
                Arguments.of("Case Qty", "Case Qty", "12"),
                Arguments.of("Cases", "Case Qty", "12"),
                Arguments.of("Case Quantity", "Case Qty", "12"),

                // Used LBS
                Arguments.of("Used LBS", "Used LBS", "123.45"),
                Arguments.of("Used Lbs.", "Used LBS", "123.45"),
                Arguments.of("LBS Used", "Used LBS", "123.45")
        );
    }

    @ParameterizedTest
    @MethodSource("headerVariants")
    void normalizesEachUploadHeaderVariant(String rawHeader, String expectedCanonical, String value) throws Exception {
        Path inDir = tempDir.resolve("input_variants");
        Files.createDirectories(inDir);

        Path out = tempDir.resolve("variants.csv");
        writeXlsx(inDir.resolve("variant.xlsx"),
                new String[]{rawHeader},
                new String[][]{{value}});

        MPRCombiner.main(new String[]{
                "--input", inDir.toString(),
                "--output", out.toString()
        });

        List<String> lines = readAllLinesUtf8(out);
        String header = lines.get(0);
        assertTrue(header.contains(expectedCanonical));
        String csv = String.join("\n", lines);
        assertTrue(csv.contains(value));
    }

    @Test
    void canWriteXlsxOutput() throws Exception {
        Path inDir = tempDir.resolve("input_xlsx_out");
        Path out = tempDir.resolve("combined.xlsx");
        Files.createDirectories(inDir);

        writeXlsx(inDir.resolve("a.xlsx"), new String[]{"ColA"}, new String[][]{{"V1"}});

        MPRCombiner.main(new String[]{
                "--input", inDir.toString(),
                "--output", out.toString()
        });

        try (var wb = new XSSFWorkbook(Files.newInputStream(out))) {
            var sheet = wb.getSheetAt(0);
            assertTrue(sheet.getRow(0).getCell(0).getStringCellValue().contains("ColA"));
            assertTrue(sheet.getRow(1).getCell(0).getStringCellValue().contains("V1"));
        }
    }

    private static void writeXlsx(Path path, String[] header, String[][] rows) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Usage");

            var headerRow = sheet.createRow(0);
            for (int c = 0; c < header.length; c++) {
                headerRow.createCell(c).setCellValue(header[c]);
            }

            for (int r = 0; r < rows.length; r++) {
                var row = sheet.createRow(r + 1);
                for (int c = 0; c < rows[r].length; c++) {
                    row.createCell(c).setCellValue(rows[r][c]);
                }
            }

            try (OutputStream os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }

    private static List<String> readAllLinesUtf8(Path p) throws IOException {
        try (BufferedReader r = Files.newBufferedReader(p, StandardCharsets.UTF_8)) {
            return r.lines().toList();
        }
    }
}

