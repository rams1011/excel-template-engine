package com.tx.excel.template;

import java.util.regex.Pattern;

final class PlaceholderPatterns {

    static final Pattern SCALAR = Pattern.compile(
            "\\{\\{\\s*([A-Za-z_][A-Za-z0-9_]*(?:\\.[A-Za-z_][A-Za-z0-9_]*)*)\\s*}}");

    private PlaceholderPatterns() {
    }
}
