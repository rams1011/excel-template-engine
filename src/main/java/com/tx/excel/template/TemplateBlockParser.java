package com.tx.excel.template;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.List;

final class TemplateBlockParser {

    private TemplateBlockParser() {
    }

    static List<TemplateBlock> parse(Sheet sheet) {
        if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
            return List.of();
        }
        return parse(sheet, sheet.getFirstRowNum(), sheet.getLastRowNum());
    }

    static List<TemplateBlock> parse(Sheet sheet, int fromRow, int toRow) {
        List<TemplateMarker> markers = scan(sheet, fromRow, toRow);
        return tree(markers);
    }

    static TemplateMarker markerInRow(Row row) {
        if (row == null) {
            return null;
        }
        for (Cell cell : row) {
            TemplateMarker marker = TemplateMarker.parse(
                    ExcelTemplateEngine.cellText(cell), row.getRowNum(), cell.getColumnIndex());
            if (marker != null) {
                return marker;
            }
        }
        return null;
    }

    private static List<TemplateMarker> scan(Sheet sheet, int fromRow, int toRow) {
        List<TemplateMarker> markers = new ArrayList<>();
        if (sheet == null || toRow < fromRow) {
            return markers;
        }
        for (int r = fromRow; r <= toRow; r++) {
            TemplateMarker marker = markerInRow(sheet.getRow(r));
            if (marker != null) {
                markers.add(marker);
            }
        }
        return markers;
    }

    private static List<TemplateBlock> tree(List<TemplateMarker> markers) {
        List<TemplateBlock> roots = new ArrayList<>();
        Deque<Frame> stack = new ArrayDeque<>();
        for (TemplateMarker marker : markers) {
            if (marker.isOpen()) {
                stack.push(new Frame(marker));
                continue;
            }
            if (stack.isEmpty()) {
                throw new ExcelTemplateException("Unexpected " + closerLabel(marker) + " at " + cellName(marker));
            }
            Frame frame = stack.pop();
            if (frame.marker.isEach() != marker.isEach()) {
                throw new ExcelTemplateException(
                        "Unbalanced " + openerLabel(frame.marker) + " at " + cellName(frame.marker)
                                + " closed by " + closerLabel(marker) + " at " + cellName(marker));
            }
            TemplateBlock block = new TemplateBlock(
                    frame.marker.isEach() ? TemplateBlock.Type.EACH : TemplateBlock.Type.IF,
                    frame.marker.path(),
                    frame.marker.row(),
                    frame.marker.column(),
                    marker.row(),
                    marker.column(),
                    frame.children);
            if (stack.isEmpty()) {
                roots.add(block);
            } else {
                stack.peek().children.add(block);
            }
        }
        if (!stack.isEmpty()) {
            TemplateMarker open = stack.peek().marker;
            throw new ExcelTemplateException("Unclosed " + openerLabel(open) + " at " + cellName(open));
        }
        return roots;
    }

    private static String openerLabel(TemplateMarker marker) {
        return marker.isEach() ? "{{#each}}" : "{{#if}}";
    }

    private static String closerLabel(TemplateMarker marker) {
        return marker.isEach() ? "{{/each}}" : "{{/if}}";
    }

    private static String cellName(TemplateMarker marker) {
        return CellReference.convertNumToColString(marker.column()) + (marker.row() + 1);
    }

    private static final class Frame {
        private final TemplateMarker marker;
        private final List<TemplateBlock> children = new ArrayList<>();

        private Frame(TemplateMarker marker) {
            this.marker = marker;
        }
    }
}
