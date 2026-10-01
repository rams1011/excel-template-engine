package com.tx.excel.template;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.net.URL;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

class ExcelTemplateEngineTest {

    private static final DataFormatter FORMATTER = new DataFormatter();

    @TempDir
    Path tempDir;

    @Test
    void replacesHeaderFooterAndTableRowsFromExcelFile() throws Exception {
        Map<String, Object> data = new HashMap<>();
        data.put("title", "Payroll");
        data.put("company", Map.of("name", "Acme"));
        data.put("period", "Sep 2026");
        data.put("printedBy", "Ram");
        data.put("total", 1250.5);
        data.put("items", List.of(
                Map.of("name", "Ada", "amount", 700),
                Map.of("name", " Bob ", "amount", 550.5)
        ));

        Path output = tempDir.resolve("payroll.xlsx");
        ExcelTemplateEngine.render(sampleTemplate(), output, data);

        try (Workbook workbook = WorkbookFactory.create(output.toFile())) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("Acme", sheet.getHeader().getLeft());
            assertEquals("Payroll", sheet.getHeader().getCenter());
            assertEquals("Ram", sheet.getFooter().getRight());
            assertEquals("Payroll - Sep 2026", cell(sheet, 0, 0));
            assertEquals("Ada", cell(sheet, 2, 0));
            assertEquals("700", cell(sheet, 2, 1));
            assertEquals(" Bob ", cell(sheet, 3, 0));
            assertEquals("550.5", cell(sheet, 3, 1));
            assertEquals("Total", cell(sheet, 4, 0));
            assertEquals("1250.5", cell(sheet, 4, 1));
        }
    }

    @Test
    void emptyTableClearsPrototypePlaceholdersFromExcelFile() throws Exception {
        Map<String, Object> data = new HashMap<>();
        data.put("title", "Empty");
        data.put("company", Map.of("name", "Acme"));
        data.put("period", "Sep 2026");
        data.put("printedBy", "Ram");
        data.put("total", 0);
        data.put("items", List.of());

        Path output = tempDir.resolve("empty.xlsx");
        ExcelTemplateEngine.render(sampleTemplate(), output, data);

        try (Workbook workbook = WorkbookFactory.create(output.toFile())) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("", cell(sheet, 2, 0));
            assertEquals("", cell(sheet, 2, 1));
            assertEquals("Total", cell(sheet, 3, 0));
        }
    }

    private static Path sampleTemplate() throws Exception {
        URL resource = ExcelTemplateEngineTest.class.getResource("/templates/sample-report.xlsx");
        assertNotNull(resource, "sample-report.xlsx must be on the test classpath");
        return Path.of(resource.toURI());
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
