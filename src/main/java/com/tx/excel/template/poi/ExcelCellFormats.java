package com.tx.excel.template.poi;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

public final class ExcelCellFormats {

    private ExcelCellFormats() {
    }

    public static boolean isText(CellStyle style) {
        if (style == null) {
            return false;
        }
        if (style.getDataFormat() == 49) {
            return true;
        }
        String format = style.getDataFormatString();
        if (format == null) {
            return false;
        }
        String trimmed = format.trim();
        return "@".equals(trimmed)
                || "text".equalsIgnoreCase(trimmed)
                || BuiltinFormats.getBuiltinFormat(49).equals(trimmed);
    }

    public static boolean isDate(CellStyle style) {
        if (style == null) {
            return false;
        }
        return DateUtil.isADateFormat(style.getDataFormat(), style.getDataFormatString());
    }

    public static boolean isNumeric(CellStyle style) {
        if (style == null || isText(style) || isDate(style)) {
            return false;
        }
        String format = style.getDataFormatString();
        if (format == null || "General".equalsIgnoreCase(format)) {
            return false;
        }
        return format.contains("0") || format.contains("#");
    }

    public static String stringify(Object value) {
        if (value == null) {
            return "";
        }
        if (value instanceof Date date) {
            return DateUtil.getExcelDate(date) + "";
        }
        return String.valueOf(value);
    }

    public static Date toDate(Object value) {
        if (value instanceof Date date) {
            return date;
        }
        if (value instanceof LocalDate date) {
            return Date.from(date.atStartOfDay(ZoneId.systemDefault()).toInstant());
        }
        if (value instanceof LocalDateTime dateTime) {
            return Date.from(dateTime.atZone(ZoneId.systemDefault()).toInstant());
        }
        if (value instanceof Number number) {
            return DateUtil.getJavaDate(number.doubleValue());
        }
        return null;
    }

    public static Double toNumber(Object value) {
        if (value instanceof Number number) {
            return number.doubleValue();
        }
        if (value instanceof String text) {
            String cleaned = text.trim().replace("$", "").replace(",", "");
            if (cleaned.isEmpty()) {
                return null;
            }
            try {
                return Double.valueOf(cleaned);
            } catch (NumberFormatException ignored) {
                return null;
            }
        }
        return null;
    }

    public static Object readTyped(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellStyle style = cell.getCellStyle();
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell) || isDate(style)) {
                    yield cell.getDateCellValue();
                }
                yield cell.getNumericCellValue();
            }
            case FORMULA -> cell.getCellFormula();
            default -> null;
        };
    }
}
