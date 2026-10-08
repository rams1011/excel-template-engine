package com.tx.excel.template;

import java.util.regex.Pattern;

/**
 * Shared placeholder syntax for the engine. Block open/close markers ({@code #each}, {@code #if}) are
 * parsed separately in {@link com.tx.excel.template.block.TemplateMarker}; this pattern only
 * matches scalar paths like {@code {{company.name}}}.
 */
final class PlaceholderPatterns {

    /** Capturing group 1 is the dot-separated path (letters, digits, underscore). */
    static final Pattern SCALAR = Pattern.compile(
            "\\{\\{\\s*([A-Za-z_][A-Za-z0-9_]*(?:\\.[A-Za-z_][A-Za-z0-9_]*)*)\\s*}}");

    private PlaceholderPatterns() {
    }
}
