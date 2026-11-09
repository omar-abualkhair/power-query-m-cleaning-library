# Power Query M Cleaning Library

A reusable set of Power Query (M) functions for cleaning raw CSV data in **Power BI** and **Excel**.
Paste it once as a query named `Clean`, then call any function from any other query:

```m
Clean[Standard](Clean[LoadCsv]("C:\Data\sales.csv"))
```

I built it so that every new project starts from the same tested cleaning steps instead of re-clicking the same transformations in the UI.

## Install (2 minutes)

1. Power Query Editor → **New Source → Blank Query**
2. **Advanced Editor** → paste the whole content of [`Clean.pq`](Clean.pq) → **Done**
3. Rename the query to **`Clean`**
4. Right-click the query → uncheck **Enable load** (it is a function library, not a table)

## Functions

| Function | What it does |
|---|---|
| `LoadCsv(filePath, delimiter?, skipLines?, encoding?)` | Loads a CSV, skips junk lines **before** the header (e.g. Google Trends exports need `2`), promotes headers. UTF-8 by default so Arabic text stays readable. |
| `TidyHeaders(tbl)` | Trims spaces and hidden characters from column names. |
| `CleanValues(tbl, nullTokens?)` | In every text value: replaces non-breaking spaces, removes hidden characters, trims, and turns `""` plus your tokens (e.g. `{"N/A", "-"}`) into `null`. |
| `AutoTypes(tbl, culture?, keepAsText?)` | Detects each column's type from **all** its values: whole numbers → Int64, decimals → number, otherwise text. `keepAsText` protects IDs and phone numbers (leading zeros). |
| `ParseDates(tbl, columns, formats, culture?)` | Parses text dates with explicit formats such as `{"dd/MM/yyyy", "yyyy-MM-dd"}`. Unknown formats raise an error instead of silently guessing. |
| `MapValues(tbl, columns, mapping)` | Replaces values using a record, e.g. `[#"Burma" = "Myanmar"]`. |
| `StandardizeCase(tbl, columns, mode?)` | `"lower"` (default), `"upper"` or `"proper"`: makes `Paid Search` / `PAID SEARCH` one category. |
| `PercentToNumber(tbl, columns, asFraction?)` | `"61%"` → `0.61` (or `61`). Values that are already numbers are kept. |
| `CountDuplicates(tbl, keyColumns?)` | Counts duplicate rows so you can log them **before** deciding to remove them. |
| `RemoveDuplicates(tbl, keyColumns?)` | Removes full-row or key-based duplicates. |
| `AddBucket(tbl, sourceColumn, newColumn, upperBounds, labels, noDataLabel?)` | Adds a category column (e.g. GDP bands) **plus** an `<newColumn> Order` column for correct sorting in Power BI. |
| `WideToLong(tbl, idColumns, attributeName, valueName)` | Unpivots one-column-per-item tables into one row per item. Note: unpivot drops null cells. |
| `Standard(tbl, culture?, keepAsText?, nullTokens?)` | The safe default pipeline: `TidyHeaders → CleanValues → remove empty rows → AutoTypes`. It does **not** remove duplicates on purpose. |

## Example

```m
let
    Source  = Clean[LoadCsv]("C:\Data\interest_over_time_weekly.csv", ",", 2),
    Cleaned = Clean[Standard](Source),
    Dated   = Clean[ParseDates](Cleaned, {"Week"}, {"yyyy-MM-dd"}),
    Long    = Clean[WideToLong](Dated, {"Week"}, "Brand", "Interest")
in
    Long
```

More examples are at the bottom of [`Clean.pq`](Clean.pq).

## Design choices

- **Explicit over clever:** dates need a format you name, and nothing is deleted silently.
- **Log before you delete:** duplicates are counted first, because identical rows are not always errors.
- **One library, many projects:** custom functions are stored in one record, so each function is called as `Clean[Name]`.

## Limitations

- `AutoTypes` scans every value, so it is slow on very large tables (1M+ rows). Set types manually there.
- Written for CSV sources. With database sources, prefer steps that keep **query folding**.

## Author

**Omar Saad Abu Al-Khair**, Data Analyst · [LinkedIn](https://www.linkedin.com/in/omar-abualkhair)
