package com.tx.excel.template;

import java.util.ArrayList;
import java.util.List;

final class TemplateBlock {

    enum Type {
        EACH,
        IF
    }

    private final Type type;
    private final String path;
    private final int openRow;
    private final int openColumn;
    private final int closeRow;
    private final int closeColumn;
    private final List<TemplateBlock> children;

    TemplateBlock(
            Type type,
            String path,
            int openRow,
            int openColumn,
            int closeRow,
            int closeColumn,
            List<TemplateBlock> children) {
        this.type = type;
        this.path = path;
        this.openRow = openRow;
        this.openColumn = openColumn;
        this.closeRow = closeRow;
        this.closeColumn = closeColumn;
        this.children = children == null ? List.of() : List.copyOf(children);
    }

    Type type() {
        return type;
    }

    String path() {
        return path;
    }

    int openRow() {
        return openRow;
    }

    int openColumn() {
        return openColumn;
    }

    int closeRow() {
        return closeRow;
    }

    int closeColumn() {
        return closeColumn;
    }

    List<TemplateBlock> children() {
        return children;
    }

    int bodyStart() {
        return openRow + 1;
    }

    int bodyEnd() {
        return closeRow - 1;
    }

    int bodyHeight() {
        return Math.max(0, closeRow - openRow - 1);
    }

    TemplateBlock shifted(int rowOffset) {
        if (rowOffset == 0) {
            return this;
        }
        List<TemplateBlock> shifted = new ArrayList<>(children.size());
        for (TemplateBlock child : children) {
            shifted.add(child.shifted(rowOffset));
        }
        return new TemplateBlock(type, path, openRow + rowOffset, openColumn,
                closeRow + rowOffset, closeColumn, shifted);
    }
}
