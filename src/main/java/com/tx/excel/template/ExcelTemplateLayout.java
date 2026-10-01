package com.tx.excel.template;

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

    public ExcelTemplateLayout header(int row, int column, String key) {
        headers.add(new HeaderCell(row, column, key));
        return this;
    }

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

    record HeaderCell(int row, int column, String key) {
    }
}
