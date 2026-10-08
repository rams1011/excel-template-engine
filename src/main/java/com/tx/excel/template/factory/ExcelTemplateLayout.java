package com.tx.excel.template.factory;

import java.util.ArrayList;
import java.util.List;

/**
 * Describes which cells are scalar placeholders and where the repeating table starts.
 */
public final class ExcelTemplateLayout {

    private final List<HeaderCell> headers = new ArrayList<>();
    private int columnHeaderRow = -1;
    private int firstDataRow = -1;
    private String tableKey = "rows";

    /** Registers a scalar cell that will become {@code {{key}}} in the generated template. */
    public ExcelTemplateLayout header(int row, int column, String key) {
        headers.add(new HeaderCell(row, column, key));
        return this;
    }

    /**
     * Defines a repeating table: {@code columnHeaderRow} supplies human labels (converted to camelCase
     * field keys), {@code firstDataRow} is the single prototype row kept in the template.
     */
    public ExcelTemplateLayout table(int columnHeaderRow, int firstDataRow, String tableKey) {
        this.columnHeaderRow = columnHeaderRow;
        this.firstDataRow = firstDataRow;
        this.tableKey = tableKey;
        return this;
    }

    List<HeaderCell> headers() {
        return headers;
    }

    int columnHeaderRow() {
        return columnHeaderRow;
    }

    int firstDataRow() {
        return firstDataRow;
    }

    String tableKey() {
        return tableKey;
    }

    public record HeaderCell(int row, int column, String key) {
    }
}
