package com.tx.excel.template.data;

import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public final class ExcelTemplateData {

    private final Map<String, Object> values;

    public ExcelTemplateData(Map<String, ?> values) {
        this.values = values == null ? Map.of() : new LinkedHashMap<>(values);
    }

    public Object get(String path) {
        if (path == null || path.isBlank()) {
            return null;
        }
        Object current = values;
        for (String segment : path.split("\\.")) {
            current = child(current, segment);
            if (current == null) {
                return null;
            }
        }
        return current;
    }

    public boolean isTable(String path) {
        String root = rootKey(path);
        return root != null && isList(values.get(root));
    }

    public String tableKey(String path) {
        return isTable(path) ? rootKey(path) : null;
    }

    public List<?> tableRows(String listKey) {
        return listAt(listKey);
    }

    public List<?> listAt(String path) {
        return asList(get(path));
    }

    public boolean isTruthy(String path) {
        return isTruthyValue(get(path));
    }

    public static boolean isTruthyValue(Object value) {
        if (value == null) {
            return false;
        }
        if (value instanceof Boolean flag) {
            return flag;
        }
        if (value instanceof Number number) {
            return number.doubleValue() != 0;
        }
        if (isList(value)) {
            return !asList(value).isEmpty();
        }
        String text = String.valueOf(value).trim();
        if (text.isEmpty() || "false".equalsIgnoreCase(text) || "0".equals(text)) {
            return false;
        }
        return true;
    }

    public ExcelTemplateData forTableRow(String listKey, Object item) {
        Map<String, Object> next = new LinkedHashMap<>(values);
        next.put(listKey, item);
        if (item instanceof Map<?, ?> map) {
            map.forEach((key, value) -> {
                if (key != null) {
                    next.put(String.valueOf(key), value);
                }
            });
        }
        return new ExcelTemplateData(next);
    }

    private static Object child(Object current, String key) {
        if (current instanceof Map<?, ?> map) {
            if (map.containsKey(key)) {
                return map.get(key);
            }
            for (Map.Entry<?, ?> entry : map.entrySet()) {
                if (entry.getKey() != null && key.equalsIgnoreCase(String.valueOf(entry.getKey()))) {
                    return entry.getValue();
                }
            }
        }
        return null;
    }

    private static String rootKey(String path) {
        if (path == null || path.isBlank()) {
            return null;
        }
        int dot = path.indexOf('.');
        return dot < 0 ? path : path.substring(0, dot);
    }

    public static boolean isList(Object value) {
        return value instanceof Collection<?> || value instanceof Object[];
    }

    private static List<?> asList(Object value) {
        if (value instanceof Collection<?> collection) {
            return new ArrayList<>(collection);
        }
        if (value instanceof Object[] array) {
            return List.of(array);
        }
        return List.of();
    }
}
