package com.tx.excel.template;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Fills an .xlsx template from a {@link Map}.
 *
 * <p>Scalar placeholders use {@code {{key}}} or {@code {{nested.key}}}. Visible block markers
 * {@code {{#each path}}} / {@code {{/each}}} and {@code {{#if path}}} / {@code {{/if}}} expand
 * or hide the rows between them (the area). A row that references a list key such as
 * {@code {{items.name}}} and is not inside {@code {{#each}}} is still treated as a single-row
 * table prototype. Excel print header and footer left/center/right are replaced the same way.</p>
 */
public final class ExcelTemplateEngine {

    private static final Pattern PLACEHOLDER = Pattern.compile(
            "\\{\\{\\s*([A-Za-z_][A-Za-z0-9_]*(?:\\.[A-Za-z_][A-Za-z0-9_]*)*)\\s*}}");
    private static final DataFormatter FORMATTER = new DataFormatter();

    private ExcelTemplateEngine() {
    }

    public static void render(Path template, Path output, Map<String, ?> data) {
        if (template == null || output == null) {
            throw new ExcelTemplateException("Excel template and output paths are required");
        }
        try {
            Path parent = output.toAbsolutePath().getParent();
            if (parent != null) {
                Files.createDirectories(parent);
            }
            try (InputStream in = Files.newInputStream(template);
                 OutputStream out = Files.newOutputStream(output)) {
                render(in, out, data);
            }
        } catch (IOException exception) {
            throw new ExcelTemplateException("Failed to render Excel template: " + template, exception);
        }
    }

    public static byte[] render(InputStream template, Map<String, ?> data) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        render(template, out, data);
        return out.toByteArray();
    }

    public static void render(InputStream template, OutputStream output, Map<String, ?> data) {
        if (template == null || output == null) {
            throw new ExcelTemplateException("Excel template and output streams are required");
        }
        ExcelTemplateData context = new ExcelTemplateData(data);
        try (Workbook workbook = WorkbookFactory.create(template)) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                fillSheet(workbook.getSheetAt(i), context);
            }
            workbook.write(output);
        } catch (IOException exception) {
            throw new ExcelTemplateException("Failed to render Excel template", exception);
        }
    }

    private static void fillSheet(Sheet sheet, ExcelTemplateData data) {
        replaceHeaderFooter(sheet, data);
        TemplateBlockExpander.expand(sheet, data);
        List<TableRow> tables = findTableRows(sheet, data);
        for (int i = tables.size() - 1; i >= 0; i--) {
            expandTable(sheet, tables.get(i), data);
        }
        for (Row row : sheet) {
            fillRow(row, data);
        }
    }

    private static void replaceHeaderFooter(Sheet sheet, ExcelTemplateData data) {
        replaceHeader(sheet.getHeader(), data);
        replaceFooter(sheet.getFooter(), data);
        if (sheet instanceof XSSFSheet xssf) {
            replaceHeader(xssf.getEvenHeader(), data);
            replaceHeader(xssf.getFirstHeader(), data);
            replaceFooter(xssf.getEvenFooter(), data);
            replaceFooter(xssf.getFirstFooter(), data);
        }
    }

    private static void replaceHeader(Header header, ExcelTemplateData data) {
        if (header == null) {
            return;
        }
        header.setLeft(replace(header.getLeft(), data));
        header.setCenter(replace(header.getCenter(), data));
        header.setRight(replace(header.getRight(), data));
    }

    private static void replaceFooter(Footer footer, ExcelTemplateData data) {
        if (footer == null) {
            return;
        }
        footer.setLeft(replace(footer.getLeft(), data));
        footer.setCenter(replace(footer.getCenter(), data));
        footer.setRight(replace(footer.getRight(), data));
    }

    private static List<TableRow> findTableRows(Sheet sheet, ExcelTemplateData data) {
        List<TableRow> tables = new ArrayList<>();
        for (Row row : sheet) {
            String listKey = tableKey(row, data);
            if (listKey != null) {
                tables.add(new TableRow(row.getRowNum(), listKey));
            }
        }
        return tables;
    }

    private static String tableKey(Row row, ExcelTemplateData data) {
        if (TemplateBlockParser.markerInRow(row) != null) {
            return null;
        }
        for (Cell cell : row) {
            String text = cellText(cell);
            if (text == null) {
                continue;
            }
            Matcher matcher = PLACEHOLDER.matcher(text);
            while (matcher.find()) {
                String key = data.tableKey(matcher.group(1));
                if (key != null) {
                    return key;
                }
            }
        }
        return null;
    }

    private static void expandTable(Sheet sheet, TableRow table, ExcelTemplateData data) {
        List<?> items = data.tableRows(table.listKey);
        Row prototype = sheet.getRow(table.rowIndex);
        if (prototype == null) {
            return;
        }
        if (items.isEmpty()) {
            fillRow(prototype, data.forTableRow(table.listKey, Map.of()));
            return;
        }
        int extra = items.size() - 1;
        if (extra > 0) {
            int last = sheet.getLastRowNum();
            if (table.rowIndex < last) {
                sheet.shiftRows(table.rowIndex + 1, last, extra, true, false);
            }
            List<CellRangeAddress> merges = rowMerges(sheet, table.rowIndex);
            for (int i = 1; i < items.size(); i++) {
                int dest = table.rowIndex + i;
                copyRow(sheet, prototype, dest);
                copyMerges(sheet, merges, dest);
            }
        }
        for (int i = 0; i < items.size(); i++) {
            Row row = sheet.getRow(table.rowIndex + i);
            fillRow(row, data.forTableRow(table.listKey, items.get(i)));
        }
    }

    static List<CellRangeAddress> rowMerges(Sheet sheet, int rowIndex) {
        List<CellRangeAddress> merges = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() == rowIndex && region.getLastRow() == rowIndex) {
                merges.add(region);
            }
        }
        return merges;
    }

    static void copyMerges(Sheet sheet, List<CellRangeAddress> merges, int destRow) {
        for (CellRangeAddress region : merges) {
            sheet.addMergedRegion(new CellRangeAddress(
                    destRow, destRow, region.getFirstColumn(), region.getLastColumn()));
        }
    }

    static void copyRow(Sheet sheet, Row source, int destIndex) {
        Row dest = sheet.getRow(destIndex);
        if (dest == null) {
            dest = sheet.createRow(destIndex);
        }
        dest.setHeight(source.getHeight());
        short first = source.getFirstCellNum();
        short last = source.getLastCellNum();
        if (first < 0) {
            return;
        }
        for (int i = first; i < last; i++) {
            Cell src = source.getCell(i);
            if (src == null) {
                continue;
            }
            Cell copy = dest.getCell(i);
            if (copy == null) {
                copy = dest.createCell(i);
            }
            CellStyle style = src.getCellStyle();
            switch (src.getCellType()) {
                case STRING -> copy.setCellValue(src.getStringCellValue());
                case NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(src)) {
                        copy.setCellValue(src.getDateCellValue());
                    } else {
                        copy.setCellValue(src.getNumericCellValue());
                    }
                }
                case BOOLEAN -> copy.setCellValue(src.getBooleanCellValue());
                case FORMULA -> copy.setCellFormula(src.getCellFormula());
                default -> copy.setBlank();
            }
            copy.setCellStyle(style);
        }
    }

    static void fillRow(Row row, ExcelTemplateData data) {
        if (row == null) {
            return;
        }
        for (Cell cell : row) {
            fillCell(cell, data);
        }
    }

    private static void fillCell(Cell cell, ExcelTemplateData data) {
        if (cell == null) {
            return;
        }
        if (cell.getCellType() == CellType.FORMULA) {
            String formula = cell.getCellFormula();
            if (formula != null && formula.contains("{{")) {
                cell.setCellFormula(replace(formula, data));
            }
            return;
        }
        String text = cellText(cell);
        if (text == null || !text.contains("{{")) {
            return;
        }
        Matcher single = PLACEHOLDER.matcher(text);
        if (single.matches()) {
            setCellValue(cell, data.get(single.group(1)));
            return;
        }
        cell.setCellValue(replace(text, data));
    }

    private static void setCellValue(Cell cell, Object value) {
        CellStyle style = cell.getCellStyle();
        writeValue(cell, value, style);
        cell.setCellStyle(style);
    }

    private static void writeValue(Cell cell, Object value, CellStyle style) {
        if (value == null || ExcelTemplateData.isList(value)) {
            cell.setBlank();
            return;
        }
        if (ExcelCellFormats.isText(style)) {
            cell.setCellValue(ExcelCellFormats.stringify(value));
            return;
        }
        if (ExcelCellFormats.isDate(style)) {
            Date date = ExcelCellFormats.toDate(value);
            if (date != null) {
                cell.setCellValue(date);
            } else {
                cell.setCellValue(ExcelCellFormats.stringify(value));
            }
            return;
        }
        if (ExcelCellFormats.isNumeric(style)) {
            Double number = ExcelCellFormats.toNumber(value);
            if (number != null) {
                cell.setCellValue(number);
                return;
            }
        }
        if (value instanceof Number number) {
            cell.setCellValue(number.doubleValue());
            return;
        }
        if (value instanceof Boolean flag) {
            cell.setCellValue(flag);
            return;
        }
        if (value instanceof Date date) {
            cell.setCellValue(date);
            return;
        }
        if (value instanceof LocalDate date) {
            cell.setCellValue(date);
            return;
        }
        if (value instanceof LocalDateTime dateTime) {
            cell.setCellValue(dateTime);
            return;
        }
        cell.setCellValue(String.valueOf(value));
    }

    static String replace(String text, ExcelTemplateData data) {
        if (text == null || text.isEmpty() || !text.contains("{{")) {
            return text;
        }
        Matcher matcher = PLACEHOLDER.matcher(text);
        StringBuilder out = new StringBuilder();
        while (matcher.find()) {
            Object value = data.get(matcher.group(1));
            String replacement = value == null || ExcelTemplateData.isList(value) ? "" : String.valueOf(value);
            matcher.appendReplacement(out, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(out);
        return out.toString();
    }

    static String cellText(Cell cell) {
        if (cell == null) {
            return null;
        }
        if (cell.getCellType() == CellType.FORMULA) {
            return cell.getCellFormula();
        }
        try {
            String rich = cell.getRichStringCellValue() == null
                    ? null
                    : cell.getRichStringCellValue().getString();
            if (rich != null && !rich.isEmpty()) {
                return rich;
            }
        } catch (IllegalStateException ignored) {
            // not a string cell
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        String formatted = FORMATTER.formatCellValue(cell);
        return formatted == null || formatted.isEmpty() ? null : formatted;
    }

    private record TableRow(int rowIndex, String listKey) {
    }
}
