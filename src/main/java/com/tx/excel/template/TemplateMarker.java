package com.tx.excel.template;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * A visible block command occupying a whole cell: {@code {{#each path}}}, {@code {{#if path}}},
 * {@code {{/each}}}, or {@code {{/if}}}.
 */
record TemplateMarker(Kind kind, String path, int row, int column) {

    private static final Pattern OPEN = Pattern.compile(
            "^\\{\\{\\s*#(each|if)\\s+([A-Za-z_][A-Za-z0-9_]*(?:\\.[A-Za-z_][A-Za-z0-9_]*)*)\\s*}}$");
    private static final Pattern CLOSE = Pattern.compile("^\\{\\{\\s*/(each|if)\\s*}}$");

    enum Kind {
        EACH_OPEN,
        IF_OPEN,
        EACH_CLOSE,
        IF_CLOSE
    }

    boolean isOpen() {
        return kind == Kind.EACH_OPEN || kind == Kind.IF_OPEN;
    }

    boolean isEach() {
        return kind == Kind.EACH_OPEN || kind == Kind.EACH_CLOSE;
    }

    static TemplateMarker parse(String text, int row, int column) {
        if (text == null) {
            return null;
        }
        String trimmed = text.trim();
        Matcher open = OPEN.matcher(trimmed);
        if (open.matches()) {
            Kind kind = "each".equals(open.group(1)) ? Kind.EACH_OPEN : Kind.IF_OPEN;
            return new TemplateMarker(kind, open.group(2), row, column);
        }
        Matcher close = CLOSE.matcher(trimmed);
        if (close.matches()) {
            Kind kind = "each".equals(close.group(1)) ? Kind.EACH_CLOSE : Kind.IF_CLOSE;
            return new TemplateMarker(kind, "", row, column);
        }
        return null;
    }
}
