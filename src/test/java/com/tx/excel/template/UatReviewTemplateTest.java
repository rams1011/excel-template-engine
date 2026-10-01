package com.tx.excel.template;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.net.URL;
import java.nio.file.Path;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

class UatReviewTemplateTest {

    private static final DataFormatter FORMATTER = new DataFormatter();

    @TempDir
    Path tempDir;

    @Test
    void generatesUatReviewFromTemplateBuiltFromTheSameWorkbook() throws Exception {
        Path sample = resource("/samples/ma-dia-peril-13-uat-review.xlsx");
        ExcelTemplateLayout layout = uatLayout();
        Path template = tempDir.resolve("ma-dia-peril-13-template.xlsx");
        Path output = tempDir.resolve("ma-dia-peril-13-output.xlsx");

        ExcelTemplateFactory.writeTemplate(sample, template, layout);
        Map<String, Object> data = ExcelTemplateFactory.extractData(sample, layout);
        ExcelTemplateEngine.render(template, output, data);

        try (Workbook expected = WorkbookFactory.create(sample.toFile());
             Workbook actual = WorkbookFactory.create(output.toFile())) {
            Sheet expectedSheet = expected.getSheetAt(0);
            Sheet actualSheet = actual.getSheetAt(0);
            assertEquals(expectedSheet.getSheetName(), actualSheet.getSheetName());
            assertEquals(expectedSheet.getLastRowNum(), actualSheet.getLastRowNum());
            int lastColumn = lastColumn(expectedSheet);
            for (int c = 0; c < lastColumn; c++) {
                assertEquals(expectedSheet.getColumnWidth(c), actualSheet.getColumnWidth(c), "column width " + c);
            }
            for (int r = 0; r <= expectedSheet.getLastRowNum(); r++) {
                assertRowMatches(expectedSheet.getRow(r), actualSheet.getRow(r), r, lastColumn);
            }
        }
    }

    private static ExcelTemplateLayout uatLayout() {
        return new ExcelTemplateLayout()
                .header(0, 0, "reportId")
                .header(1, 0, "title")
                .header(2, 0, "runDate")
                .header(3, 0, "period")
                .table(5, 6, "rows");
    }

    private static void assertRowMatches(Row expected, Row actual, int rowIndex, int lastColumn) {
        if (expected == null) {
            return;
        }
        assertNotNull(actual, "missing row " + rowIndex);
        assertEquals(expected.getHeight(), actual.getHeight(), "row height " + rowIndex);
        for (int c = 0; c < lastColumn; c++) {
            Cell left = expected.getCell(c);
            Cell right = actual.getCell(c);
            if (left == null) {
                continue;
            }
            assertNotNull(right, "missing cell " + coord(rowIndex, c));
            assertEquals(left.getCellType(), right.getCellType(), "type " + coord(rowIndex, c));
            assertEquals(FORMATTER.formatCellValue(left), FORMATTER.formatCellValue(right), "value " + coord(rowIndex, c));
            CellStyle leftStyle = left.getCellStyle();
            CellStyle rightStyle = right.getCellStyle();
            assertEquals(leftStyle.getAlignment(), rightStyle.getAlignment(), "align " + coord(rowIndex, c));
            assertEquals(leftStyle.getDataFormatString(), rightStyle.getDataFormatString(), "format " + coord(rowIndex, c));
        }
    }

    private static int lastColumn(Sheet sheet) {
        int last = 0;
        for (Row row : sheet) {
            if (row.getLastCellNum() > last) {
                last = row.getLastCellNum();
            }
        }
        return last;
    }

    private static String coord(int row, int column) {
        return (char) ('A' + column) + String.valueOf(row + 1);
    }

    private static Path resource(String name) throws Exception {
        URL url = UatReviewTemplateTest.class.getResource(name);
        assertNotNull(url, name + " must be on the test classpath");
        return Path.of(url.toURI());
    }
}
