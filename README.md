# Excel Template Engine

Fill **`.xlsx`** report templates from a Java **`Map<String, ?>`**. Put **`{{placeholders}}`** in cells (and in Excel print headers/footers); pass lists to repeat table rows or use **`{{#each}}` / `{{#if}}`** blocks for multi-row sections.

Requires **Java 17**. Uses Apache POI internally—you do not need POI on the classpath unless you also manipulate workbooks yourself.

## Add the dependency

Install from source if needed: `mvn install` in this repo.

```xml
<dependency>
    <groupId>com.tx.excel</groupId>
    <artifactId>excel-template-engine</artifactId>
    <version>1.0.0-SNAPSHOT</version>
</dependency>
```

## Render a template

**File to file:**

```java
import com.tx.excel.template.ExcelTemplateEngine;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;

Map<String, Object> data = Map.of(
    "title", "Payroll",
    "company", Map.of("name", "Acme"),
    "period", "Sep 2026",
    "printedBy", "Ram",
    "total", 1250.5,
    "items", List.of(
        Map.of("name", "Ada", "amount", 700),
        Map.of("name", "Bob", "amount", 550.5)
    )
);

ExcelTemplateEngine.render(
    Path.of("report-template.xlsx"),
    Path.of("report-filled.xlsx"),
    data);
```

**Streams / bytes:**

```java
ExcelTemplateEngine.render(templateInputStream, outputStream, data);

byte[] filled = ExcelTemplateEngine.render(templateInputStream, data);
```

IO and template errors throw **`com.tx.excel.template.exception.ExcelTemplateException`** (import as `ExcelTemplateException`).

## Design your template in Excel

### Scalar fields

| In the cell | In your map |
|-------------|-------------|
| `{{title}}` | `"title"` → value |
| `{{company.name}}` | nested map under `"company"` |

Map keys are matched case-insensitively. Dot paths walk nested maps.

If a cell contains **only** one placeholder, numbers, dates, and booleans are written according to the cell’s existing format. Text like `Period: {{period}}` is plain string replacement.

### Repeating table (simple)

1. Add a column header row in Excel (labels only).
2. On the **next row**, use one prototype row, e.g. `{{items.name}}` and `{{items.amount}}`.
3. Put a **`List`** (or array) in data under `items`; each element should be a **map** (or the row fields you reference).

That row is copied once per list item. An **empty list** clears placeholders on the prototype row but keeps the row.

### Row blocks (`#each` / `#if`)

For several rows per item or optional sections, put markers in **their own cells** (one marker per cell, exact spelling):

```
{{#each items}}
   … body rows …
{{/each}}

{{#if showTotal}}
   … optional rows …
{{/if}}
```

Inside `#each`, use fields on the current item (`{{name}}`, `{{amount}}`) or full paths from the root map. Marker rows are removed in the output.

**`#if` is true when:** value is non-null; boolean `true`; non-zero number; non-empty list; non-empty string (not `false` or `0`).

### Print header and footer

In Excel: **Page Layout → Header/Footer**. Use `{{company.name}}`, `{{printedBy}}`, etc. in left, center, or right (same rules as cells).

## Optional: build a template from a filled sample

If you already have a **filled** workbook and know which cells are headers vs table data, **`ExcelTemplateFactory`** can stamp `{{…}}` and leave one table prototype row:

```java
import com.tx.excel.template.factory.ExcelTemplateFactory;
import com.tx.excel.template.factory.ExcelTemplateLayout;

ExcelTemplateLayout layout = new ExcelTemplateLayout()
    .header(0, 0, "title")
    .table(1, 2, "items");   // column header row 1, first data row 2, list key "items"

ExcelTemplateFactory.writeTemplate(filledSamplePath, templatePath, layout);
```

You can also **`extractData(sample, layout)`** to pull a map from the sample (useful for round-trip tests or migrations).

## Limitations

- **`.xlsx` only** (not `.xls`).
- **`#each` / `#if` markers** must be the entire cell contents (after trim).
- Large lists expand by copying rows in the sheet—test performance on very big reports.
- Formulas are only special-cased when the formula text contains `{{`.

## Main API

| Class | Package | Use |
|-------|---------|-----|
| **`ExcelTemplateEngine`** | `com.tx.excel.template` | `render(...)` — production path |
| **`ExcelTemplateFactory`** | `com.tx.excel.template.factory` | Create templates or extract data from samples |
| **`ExcelTemplateLayout`** | `com.tx.excel.template.factory` | Describes header cells and table rows for the factory |
| **`ExcelTemplateException`** | `com.tx.excel.template.exception` | Unchecked errors from render/factory |

Internal packages (`data`, `poi`, `block`) are not part of the public API; depend only on the classes above unless you extend the engine.
