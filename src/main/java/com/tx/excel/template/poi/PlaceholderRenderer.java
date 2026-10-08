package com.tx.excel.template.poi;

import com.tx.excel.template.data.ExcelTemplateData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Low-level cell and string substitution. When a cell contains exactly one placeholder, typed values
 * are written to match the cell's existing format; mixed text uses string replacement only.
 */
public final class PlaceholderRenderer {

    private static final Pattern PLACEHOLDER = Pattern.compile(
            "\\{\\{\\s*([A-Za-z_][A-Za-z0-9_]*(?:\\.[A-Za-z_][A-Za-z0-9_]*)*)\\s*}}");
    private static final DataFormatter FORMATTER = new DataFormatter();

    private PlaceholderRenderer() {
    }

    public static void fillRow(Row row, ExcelTemplateData data) {
        if (row == null) {
            return;
        }
        for (Cell cell : row) {
            fillCell(cell, data);
        }
    }

    /** Deep-copies cell values, formulas, and styles from {@code source} into row {@code destIndex}. */
    public static void copyRow(Sheet sheet, Row source, int destIndex) {
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

    public static List<CellRangeAddress> rowMerges(Sheet sheet, int rowIndex) {
        List<CellRangeAddress> merges = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() == rowIndex && region.getLastRow() == rowIndex) {
                merges.add(region);
            }
        }
        return merges;
    }

    public static void copyMerges(Sheet sheet, List<CellRangeAddress> merges, int destRow) {
        for (CellRangeAddress region : merges) {
            sheet.addMergedRegion(new CellRangeAddress(
                    destRow, destRow, region.getFirstColumn(), region.getLastColumn()));
        }
    }

    /** Inline substitution for headers/footers and cells with multiple placeholders; lists render as empty. */
    public static String replace(String text, ExcelTemplateData data) {
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

    /**
     * Text used to find placeholders. Formulas return the formula string (not evaluated value) so
     * {@code {{…}}} inside formulas can be rewritten before Excel recalculates.
     */
    public static String cellText(Cell cell) {
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
        // Whole-cell placeholder → preserve number/date format from the template cell.
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
}
