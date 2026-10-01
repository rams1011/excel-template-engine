package com.tx.excel.template;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Builds an .xlsx template from a filled sample workbook and extracts the map used to fill it.
 */
public final class ExcelTemplateFactory {

    private ExcelTemplateFactory() {
    }

    public static void writeTemplate(Path sample, Path template, ExcelTemplateLayout layout) {
        try {
            Path parent = template.toAbsolutePath().getParent();
            if (parent != null) {
                Files.createDirectories(parent);
            }
            try (InputStream in = Files.newInputStream(sample);
                 OutputStream out = Files.newOutputStream(template)) {
                writeTemplate(in, out, layout);
            }
        } catch (IOException exception) {
            throw new ExcelTemplateException("Failed to write Excel template from " + sample, exception);
        }
    }

    public static void writeTemplate(InputStream sample, OutputStream template, ExcelTemplateLayout layout) {
        try (Workbook workbook = WorkbookFactory.create(sample)) {
            Sheet sheet = workbook.getSheetAt(0);
            applyPlaceholders(sheet, layout);
            trimTableRows(sheet, layout.firstDataRow());
            workbook.write(template);
        } catch (IOException exception) {
            throw new ExcelTemplateException("Failed to write Excel template", exception);
        }
    }

    public static Map<String, Object> extractData(Path sample, ExcelTemplateLayout layout) {
        try (InputStream in = Files.newInputStream(sample)) {
            return extractData(in, layout);
        } catch (IOException exception) {
            throw new ExcelTemplateException("Failed to read Excel sample: " + sample, exception);
        }
    }

    public static Map<String, Object> extractData(InputStream sample, ExcelTemplateLayout layout) {
        try (Workbook workbook = WorkbookFactory.create(sample)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<String, Object> data = new LinkedHashMap<>();
            for (ExcelTemplateLayout.HeaderCell header : layout.headers()) {
                data.put(header.key(), cellText(sheet, header.row(), header.column()));
            }
            if (layout.firstDataRow() >= 0) {
                data.put(layout.tableKey(), extractRows(sheet, layout));
            }
            return data;
        } catch (IOException exception) {
            throw new ExcelTemplateException("Failed to extract Excel sample data", exception);
        }
    }

    static List<String> columnKeys(Sheet sheet, int headerRow) {
        Row row = sheet.getRow(headerRow);
        List<String> keys = new ArrayList<>();
        if (row == null) {
            return keys;
        }
        for (int c = 0; c < row.getLastCellNum(); c++) {
            keys.add(columnKey(cellText(sheet, headerRow, c)));
        }
        return keys;
    }

    static String columnKey(String header) {
        String cleaned = header == null ? "" : header.replaceAll("[^A-Za-z0-9]+", " ").trim();
        if (cleaned.isEmpty()) {
            return "column";
        }
        String[] parts = cleaned.split("\\s+");
        StringBuilder key = new StringBuilder(parts[0].toLowerCase());
        for (int i = 1; i < parts.length; i++) {
            String part = parts[i];
            key.append(Character.toUpperCase(part.charAt(0)));
            if (part.length() > 1) {
                key.append(part.substring(1));
            }
        }
        return key.toString();
    }

    private static void applyPlaceholders(Sheet sheet, ExcelTemplateLayout layout) {
        for (ExcelTemplateLayout.HeaderCell header : layout.headers()) {
            setTextKeepStyle(sheet, header.row(), header.column(), "{{" + header.key() + "}}");
        }
        if (layout.firstDataRow() < 0) {
            return;
        }
        List<String> keys = columnKeys(sheet, layout.columnHeaderRow());
        Row data = sheet.getRow(layout.firstDataRow());
        if (data == null) {
            return;
        }
        for (int c = 0; c < keys.size(); c++) {
            setTextKeepStyle(sheet, layout.firstDataRow(), c, "{{" + layout.tableKey() + "." + keys.get(c) + "}}");
        }
    }

    private static void trimTableRows(Sheet sheet, int firstDataRow) {
        if (firstDataRow < 0) {
            return;
        }
        for (int r = sheet.getLastRowNum(); r > firstDataRow; r--) {
            Row row = sheet.getRow(r);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    private static List<Map<String, Object>> extractRows(Sheet sheet, ExcelTemplateLayout layout) {
        List<String> keys = columnKeys(sheet, layout.columnHeaderRow());
        List<Map<String, Object>> rows = new ArrayList<>();
        for (int r = layout.firstDataRow(); r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            Map<String, Object> item = new LinkedHashMap<>();
            boolean empty = true;
            for (int c = 0; c < keys.size(); c++) {
                Cell cell = row.getCell(c);
                Object value = ExcelCellFormats.readTyped(cell);
                item.put(keys.get(c), value);
                if (value != null && !(value instanceof String text && text.isBlank())) {
                    empty = false;
                }
            }
            if (!empty) {
                rows.add(item);
            }
        }
        return rows;
    }

    private static void setTextKeepStyle(Sheet sheet, int rowIndex, int column, String text) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(column);
        if (cell == null) {
            cell = row.createCell(column);
        }
        CellStyle style = cell.getCellStyle();
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

    private static String cellText(Sheet sheet, int rowIndex, int column) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }
        Object value = ExcelCellFormats.readTyped(cell);
        return value == null ? "" : String.valueOf(value);
    }
}
