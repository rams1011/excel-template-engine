package com.tx.excel.template.block;

import com.tx.excel.template.data.ExcelTemplateData;
import com.tx.excel.template.poi.PlaceholderRenderer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * Expands {@code #each} and {@code #if} blocks by copying or deleting rows and removing marker rows
 * from the output. Sibling blocks are processed bottom-up so row indices remain stable.
 */
public final class TemplateBlockExpander {

    private TemplateBlockExpander() {
    }

    public static void expand(Sheet sheet, ExcelTemplateData data) {
        expandChildren(sheet, TemplateBlockParser.parse(sheet), data);
    }

    /** True when the row contains a {@code #each} / {@code #if} block marker cell. */
    public static boolean rowHasBlockMarker(org.apache.poi.ss.usermodel.Row row) {
        return TemplateBlockParser.markerInRow(row) != null;
    }

    /** Returns net row count change after expanding all blocks (used when nesting inside {@code #each}). */
    static int expandChildren(Sheet sheet, List<TemplateBlock> blocks, ExcelTemplateData data) {
        int totalDelta = 0;
        for (int i = blocks.size() - 1; i >= 0; i--) {
            totalDelta += expand(sheet, blocks.get(i), data);
        }
        return totalDelta;
    }

    private static int expand(Sheet sheet, TemplateBlock block, ExcelTemplateData data) {
        if (block.type() == TemplateBlock.Type.EACH) {
            return expandEach(sheet, block, data);
        }
        return expandIf(sheet, block, data);
    }

    private static int expandEach(Sheet sheet, TemplateBlock block, ExcelTemplateData data) {
        List<?> items = data.listAt(block.path());
        int open = block.openRow();
        int close = block.closeRow();
        int bodyStart = block.bodyStart();
        int bodyEnd = block.bodyEnd();
        int height = block.bodyHeight();
        int originalSpan = close - open + 1;

        if (items.isEmpty() || height <= 0) {
            deleteRows(sheet, open, close);
            return -originalSpan;
        }

        // Remove close marker first; body rows are copied once per list item.
        deleteRows(sheet, close, close);

        int extra = items.size() - 1;
        if (extra > 0) {
            int extraRows = extra * height;
            int last = sheet.getLastRowNum();
            if (bodyEnd < last) {
                sheet.shiftRows(bodyEnd + 1, last, extraRows, true, false);
            }
            for (int i = 1; i <= extra; i++) {
                copyRowRange(sheet, bodyStart, bodyEnd, bodyStart + i * height);
            }
        }

        int childDeltas = 0;
        for (int i = items.size() - 1; i >= 0; i--) {
            int offset = i * height;
            ExcelTemplateData itemData = data.forTableRow(block.path(), items.get(i));
            int childDelta = expandChildren(sheet, shiftAll(block.children(), offset), itemData);
            int copyStart = bodyStart + offset;
            int copyEnd = copyStart + height - 1 + childDelta;
            fillRows(sheet, copyStart, copyEnd, itemData);
            childDeltas += childDelta;
        }

        deleteRows(sheet, open, open);
        return items.size() * height + childDeltas - originalSpan;
    }

    /** Keeps body rows when the path is truthy; otherwise deletes the entire block including markers. */
    private static int expandIf(Sheet sheet, TemplateBlock block, ExcelTemplateData data) {
        int open = block.openRow();
        int close = block.closeRow();
        int originalSpan = close - open + 1;
        if (!data.isTruthy(block.path())) {
            deleteRows(sheet, open, close);
            return -originalSpan;
        }
        int childDelta = expandChildren(sheet, block.children(), data);
        int newClose = close + childDelta;
        deleteRows(sheet, newClose, newClose);
        deleteRows(sheet, open, open);
        return childDelta - 2;
    }

    private static List<TemplateBlock> shiftAll(List<TemplateBlock> blocks, int rowOffset) {
        if (rowOffset == 0 || blocks.isEmpty()) {
            return blocks;
        }
        List<TemplateBlock> shifted = new ArrayList<>(blocks.size());
        for (TemplateBlock block : blocks) {
            shifted.add(block.shifted(rowOffset));
        }
        return shifted;
    }

    private static void fillRows(Sheet sheet, int fromRow, int toRow, ExcelTemplateData data) {
        for (int r = fromRow; r <= toRow; r++) {
            PlaceholderRenderer.fillRow(sheet.getRow(r), data);
        }
    }

    private static void copyRowRange(Sheet sheet, int srcStart, int srcEnd, int destStart) {
        List<CellRangeAddress> merges = mergesInRange(sheet, srcStart, srcEnd);
        int offset = destStart - srcStart;
        for (int r = srcStart; r <= srcEnd; r++) {
            Row source = sheet.getRow(r);
            if (source == null) {
                continue;
            }
            PlaceholderRenderer.copyRow(sheet, source, r + offset);
        }
        for (CellRangeAddress region : merges) {
            sheet.addMergedRegion(new CellRangeAddress(
                    region.getFirstRow() + offset,
                    region.getLastRow() + offset,
                    region.getFirstColumn(),
                    region.getLastColumn()));
        }
    }

    private static List<CellRangeAddress> mergesInRange(Sheet sheet, int firstRow, int lastRow) {
        List<CellRangeAddress> merges = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() >= firstRow && region.getLastRow() <= lastRow) {
                merges.add(region);
            }
        }
        return merges;
    }

    /** Removes merged regions overlapping the range, clears rows, then shifts rows below upward. */
    static void deleteRows(Sheet sheet, int first, int last) {
        if (first > last) {
            return;
        }
        removeMergesOverlapping(sheet, first, last);
        int count = last - first + 1;
        int lastRow = sheet.getLastRowNum();
        for (int r = first; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
        if (last < lastRow) {
            sheet.shiftRows(last + 1, lastRow, -count, true, false);
        }
    }

    private static void removeMergesOverlapping(Sheet sheet, int first, int last) {
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getLastRow() >= first && region.getFirstRow() <= last) {
                sheet.removeMergedRegion(i);
            }
        }
    }
}
