package com.tx.excel.template;

import com.tx.excel.template.block.TemplateMarker;
import com.tx.excel.template.exception.ExcelTemplateException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelTemplateBlockTest {

    private static final DataFormatter FORMATTER = new DataFormatter();

    @TempDir
    Path tempDir;

    @BeforeAll
    static void writeClasspathSample() throws Exception {
        byte[] bytes = eachIfTemplate().readAllBytes();
        Path source = Path.of("src/test/resources/templates/each-if-report.xlsx");
        Path compiled = Path.of("target/test-classes/templates/each-if-report.xlsx");
        Files.createDirectories(source.getParent());
        Files.createDirectories(compiled.getParent());
        Files.write(source, bytes);
        Files.write(compiled, bytes);
    }

    @Test
    void expandsEachIfReportFromClasspathTemplate() throws Exception {
        URL resource = ExcelTemplateBlockTest.class.getResource("/templates/each-if-report.xlsx");
        assertNotNull(resource, "each-if-report.xlsx must be on the test classpath");
        Map<String, Object> data = new HashMap<>();
        data.put("title", "Orders");
        data.put("total", 1250);
        data.put("items", List.of(
                map("name", "Ada", "amount", 700, "note", "Rush"),
                map("name", "Bob", "amount", 550)
        ));
        Path output = tempDir.resolve("from-classpath.xlsx");
        ExcelTemplateEngine.render(Path.of(resource.toURI()), output, data);
        try (Workbook workbook = WorkbookFactory.create(output.toFile())) {
            assertEquals("Orders", cell(workbook.getSheetAt(0), 0, 0));
            assertEquals("Ada", cell(workbook.getSheetAt(0), 1, 0));
            assertEquals("Bob", cell(workbook.getSheetAt(0), 3, 0));
            assertEquals("Total", cell(workbook.getSheetAt(0), 4, 0));
        }
    }

    @Test
    void expandsMultiRowEachAndNestedIfThenRemovesMarkers() throws Exception {
        Map<String, Object> data = new HashMap<>();
        data.put("title", "Orders");
        data.put("total", 1250);
        data.put("items", List.of(
                map("name", "Ada", "amount", 700, "note", "Rush"),
                map("name", "Bob", "amount", 550)
        ));

        Path output = tempDir.resolve("orders.xlsx");
        Files.write(output, ExcelTemplateEngine.render(eachIfTemplate(), data));

        try (Workbook workbook = WorkbookFactory.create(output.toFile())) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("Orders", cell(sheet, 0, 0));
            assertEquals("Ada", cell(sheet, 1, 0));
            assertEquals("700", cell(sheet, 1, 1));
            assertEquals("Note: Rush", cell(sheet, 2, 0));
            assertEquals("Bob", cell(sheet, 3, 0));
            assertEquals("550", cell(sheet, 3, 1));
            assertEquals("Total", cell(sheet, 4, 0));
            assertEquals("1250", cell(sheet, 4, 1));
            assertNoMarkers(sheet);
        }
    }

    @Test
    void falseIfRemovesTheBlockAndShiftsFollowingRowsUp() throws Exception {
        Map<String, Object> data = new HashMap<>();
        data.put("title", "Title");
        data.put("showNotes", false);
        data.put("notes", "hidden");

        byte[] rendered = ExcelTemplateEngine.render(ifOnlyTemplate(), data);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(rendered))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("Title", cell(sheet, 0, 0));
            assertEquals("Footer", cell(sheet, 1, 0));
            assertNoMarkers(sheet);
        }
    }

    @Test
    void emptyEachDropsTheWholeBlock() throws Exception {
        Map<String, Object> data = new HashMap<>();
        data.put("title", "Head");
        data.put("items", List.of());

        byte[] rendered = ExcelTemplateEngine.render(emptyEachTemplate(), data);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(rendered))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("Head", cell(sheet, 0, 0));
            assertEquals("Tail", cell(sheet, 1, 0));
            assertNoMarkers(sheet);
        }
    }

    @Test
    void unbalancedCloserFailsWithCellLocation() {
        ExcelTemplateException error = assertThrows(ExcelTemplateException.class, () ->
                ExcelTemplateEngine.render(unbalancedTemplate(), Map.of("title", "X")));
        assertTrue(error.getMessage().contains("{{/if}}"));
        assertTrue(error.getMessage().contains("A2"));
    }

    private static InputStream eachIfTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Report");
            row(sheet, 0, "{{title}}");
            row(sheet, 1, "{{#each items}}");
            row(sheet, 2, "{{name}}", "{{amount}}");
            row(sheet, 3, "{{#if note}}");
            row(sheet, 4, "Note: {{note}}");
            row(sheet, 5, "{{/if}}");
            row(sheet, 6, "{{/each}}");
            row(sheet, 7, "Total", "{{total}}");
            return bytes(workbook);
        }
    }

    private static InputStream ifOnlyTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Notes");
            row(sheet, 0, "{{title}}");
            row(sheet, 1, "{{#if showNotes}}");
            row(sheet, 2, "{{notes}}");
            row(sheet, 3, "{{/if}}");
            row(sheet, 4, "Footer");
            return bytes(workbook);
        }
    }

    private static InputStream emptyEachTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Empty");
            row(sheet, 0, "{{title}}");
            row(sheet, 1, "{{#each items}}");
            row(sheet, 2, "{{name}}");
            row(sheet, 3, "{{/each}}");
            row(sheet, 4, "Tail");
            return bytes(workbook);
        }
    }

    private static InputStream unbalancedTemplate() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Bad");
            row(sheet, 0, "{{title}}");
            row(sheet, 1, "{{/if}}");
            return bytes(workbook);
        }
    }

    private static InputStream bytes(Workbook workbook) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        return new ByteArrayInputStream(out.toByteArray());
    }

    private static void row(Sheet sheet, int index, String... values) {
        Row row = sheet.createRow(index);
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
    }

    private static Map<String, Object> map(Object... pairs) {
        Map<String, Object> values = new HashMap<>();
        for (int i = 0; i < pairs.length; i += 2) {
            values.put(String.valueOf(pairs[i]), pairs[i + 1]);
        }
        return values;
    }

    private static void assertNoMarkers(Sheet sheet) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = ExcelTemplateEngine.cellText(cell);
                if (text != null) {
                    assertTrue(TemplateMarker.parse(text, row.getRowNum(), cell.getColumnIndex()) == null,
                            "leftover marker " + text);
                }
            }
        }
    }

    private static String cell(Sheet sheet, int rowIndex, int column) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(column);
        return cell == null ? "" : FORMATTER.formatCellValue(cell);
    }
}
