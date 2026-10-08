package com.tx.excel.template;

import com.tx.excel.template.block.TemplateBlockExpander;
import com.tx.excel.template.data.ExcelTemplateData;
import com.tx.excel.template.exception.ExcelTemplateException;
import com.tx.excel.template.poi.PlaceholderRenderer;
import org.apache.poi.ss.usermodel.Cell;
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
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

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

    private ExcelTemplateEngine() {
    }

    /** Renders {@code template} to {@code output}, creating parent directories if needed. */
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

    /**
     * Reads cell text for tests and diagnostics (formulas return the formula string).
     */
    public static String cellText(Cell cell) {
        return PlaceholderRenderer.cellText(cell);
    }

    /**
     * Per-sheet pipeline: print areas first, then row blocks (which insert/delete rows), then simple
     * list tables (bottom-up so row indices stay valid), finally scalar placeholders on every row.
     */
    private static void fillSheet(Sheet sheet, ExcelTemplateData data) {
        replaceHeaderFooter(sheet, data);
        TemplateBlockExpander.expand(sheet, data);
        List<TableRow> tables = findTableRows(sheet, data);
        for (int i = tables.size() - 1; i >= 0; i--) {
            expandTable(sheet, tables.get(i), data);
        }
        for (Row row : sheet) {
            PlaceholderRenderer.fillRow(row, data);
        }
    }

    private static void replaceHeaderFooter(Sheet sheet, ExcelTemplateData data) {
        replaceHeader(sheet.getHeader(), data);
        replaceFooter(sheet.getFooter(), data);
        // XSSF supports alternate first/even page header and footer sections.
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
        header.setLeft(PlaceholderRenderer.replace(header.getLeft(), data));
        header.setCenter(PlaceholderRenderer.replace(header.getCenter(), data));
        header.setRight(PlaceholderRenderer.replace(header.getRight(), data));
    }

    private static void replaceFooter(Footer footer, ExcelTemplateData data) {
        if (footer == null) {
            return;
        }
        footer.setLeft(PlaceholderRenderer.replace(footer.getLeft(), data));
        footer.setCenter(PlaceholderRenderer.replace(footer.getCenter(), data));
        footer.setRight(PlaceholderRenderer.replace(footer.getRight(), data));
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

    /** Detects a one-row table prototype: first placeholder whose root key is a list in {@code data}. */
    private static String tableKey(Row row, ExcelTemplateData data) {
        if (TemplateBlockExpander.rowHasBlockMarker(row)) {
            return null;
        }
        for (Cell cell : row) {
            String text = PlaceholderRenderer.cellText(cell);
            if (text == null) {
                continue;
            }
            Matcher matcher = PlaceholderPatterns.SCALAR.matcher(text);
            while (matcher.find()) {
                String key = data.tableKey(matcher.group(1));
                if (key != null) {
                    return key;
                }
            }
        }
        return null;
    }

    /** Duplicates the prototype row for each list element, then fills each copy with that item's context. */
    private static void expandTable(Sheet sheet, TableRow table, ExcelTemplateData data) {
        List<?> items = data.tableRows(table.listKey);
        Row prototype = sheet.getRow(table.rowIndex);
        if (prototype == null) {
            return;
        }
        if (items.isEmpty()) {
            PlaceholderRenderer.fillRow(prototype, data.forTableRow(table.listKey, Map.of()));
            return;
        }
        int extra = items.size() - 1;
        if (extra > 0) {
            int last = sheet.getLastRowNum();
            if (table.rowIndex < last) {
                sheet.shiftRows(table.rowIndex + 1, last, extra, true, false);
            }
            List<CellRangeAddress> merges = PlaceholderRenderer.rowMerges(sheet, table.rowIndex);
            for (int i = 1; i < items.size(); i++) {
                int dest = table.rowIndex + i;
                PlaceholderRenderer.copyRow(sheet, prototype, dest);
                PlaceholderRenderer.copyMerges(sheet, merges, dest);
            }
        }
        for (int i = 0; i < items.size(); i++) {
            Row row = sheet.getRow(table.rowIndex + i);
            PlaceholderRenderer.fillRow(row, data.forTableRow(table.listKey, items.get(i)));
        }
    }

    /** Simple table: one template row bound to a top-level list key. */
    private record TableRow(int rowIndex, String listKey) {
    }
}
