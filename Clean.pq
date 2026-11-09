// =====================================================================
//  M CLEANING LIBRARY  -  reusable Power Query functions
//  HOW TO ADD:  Power Query > New Source > Blank Query > Advanced Editor
//               paste this whole file > Done > rename the query to:  Clean
//  HOW TO USE:  in any other query:   Clean[FunctionName](arguments)
// =====================================================================
let
    // ---------------------------------------------------------------
    // 1) LoadCsv: load a CSV, skip junk lines BEFORE the header, promote headers
    //    filePath  : full path to the file
    //    delimiter : "," (default) or ";" ...
    //    skipLines : number of lines to skip before the header row (default 0)
    //                Google Trends files need 2
    //    encoding  : 65001 = UTF-8 (default, keeps Arabic text correct)
    // ---------------------------------------------------------------
    fnLoadCsv = (filePath as text, optional delimiter as nullable text, optional skipLines as nullable number, optional encoding as nullable number) as table =>
        let
            delim   = if delimiter = null then "," else delimiter,
            skip    = if skipLines = null then 0 else skipLines,
            enc     = if encoding = null then 65001 else encoding,
            Lines   = Lines.FromBinary(File.Contents(filePath), QuoteStyle.Csv, false, enc),
            Kept    = List.Skip(Lines, skip),
            Parsed  = Csv.Document(Text.Combine(Kept, "#(lf)"), [Delimiter = delim, QuoteStyle = QuoteStyle.Csv]),
            Headers = Table.PromoteHeaders(Parsed, [PromoteAllScalars = true])
        in
            Headers,

    // ---------------------------------------------------------------
    // 2) TidyHeaders: trim spaces and hidden characters from column names
    // ---------------------------------------------------------------
    fnTidyHeaders = (tbl as table) as table =>
        Table.TransformColumnNames(tbl, each Text.Trim(Text.Clean(_))),

    // ---------------------------------------------------------------
    // 3) CleanValues: in EVERY column, for text values:
    //    - replace non-breaking spaces, remove hidden characters, trim
    //    - turn "" and any extra tokens you give into null
    //    nullTokens example: {"N/A", "NA", "-", "null"}
    // ---------------------------------------------------------------
    fnCleanValues = (tbl as table, optional nullTokens as nullable list) as table =>
        let
            tokens = if nullTokens = null then {""} else List.Combine({{""}, nullTokens}),
            fx = (v) =>
                if v is text then
                    let t = Text.Trim(Text.Clean(Text.Replace(v, "#(00A0)", " ")))
                    in  if List.Contains(tokens, t) then null else t
                else v
        in
            Table.TransformColumns(tbl, List.Transform(Table.ColumnNames(tbl), each {_, fx})),

    // ---------------------------------------------------------------
    // 4) AutoTypes: detect each column's type from ALL its values
    //    whole numbers -> Int64, decimals -> number, anything else -> text
    //    keepAsText: columns that must stay text even if they look numeric
    //                (IDs, phone numbers, zip codes - keeps leading zeros)
    //    NOTE: scans every value, so it is slow on very large tables (1M+ rows)
    // ---------------------------------------------------------------
    fnAutoTypes = (tbl as table, optional culture as nullable text, optional keepAsText as nullable list) as table =>
        let
            cult = if culture = null then "en-US" else culture,
            keep = if keepAsText = null then {} else keepAsText,
            cols = List.RemoveItems(Table.ColumnNames(tbl), keep),
            TypeList = List.Transform(cols, (col) =>
                let
                    values     = List.RemoveNulls(Table.Column(tbl, col)),
                    nums       = List.Transform(values, each try Number.From(_, cult) otherwise null),
                    allNumeric = List.Count(values) > 0 and not List.Contains(nums, null),
                    allWhole   = allNumeric and List.AllTrue(List.Transform(nums, each _ = Number.RoundDown(_)))
                in
                    {col, if allWhole then Int64.Type else if allNumeric then type number else type text}),
            KeepList = List.Transform(keep, each {_, type text})
        in
            Table.TransformColumnTypes(tbl, List.Combine({TypeList, KeepList}), cult),

    // ---------------------------------------------------------------
    // 5) ParseDates: dates written in different formats in the same column
    //    formats are tried IN ORDER - you decide what "05/03/2023" means
    //    example formats: {"yyyy-MM-dd", "dd/MM/yyyy", "MMM dd, yyyy"}
    //    an unknown format becomes an ERROR on purpose, so you see it
    // ---------------------------------------------------------------
    fnParseDates = (tbl as table, columns as list, formats as list, optional culture as nullable text) as table =>
        let
            cult = if culture = null then "en-US" else culture,
            parseOne = (v) =>
                if v = null then null
                else if v is date then v
                else if v is datetime then Date.From(v)
                else
                    let
                        attempts = List.Transform(formats, (f) => try Date.FromText(Text.From(v), [Format = f, Culture = cult]) otherwise null),
                        found    = List.First(List.RemoveNulls(attempts), null)
                    in
                        if found = null then error Error.Record("DateParseError", "Unrecognized date: " & Text.From(v)) else found
        in
            Table.TransformColumns(tbl, List.Transform(columns, (c) => {c, parseOne, type nullable date})),

    // ---------------------------------------------------------------
    // 6) MapValues: fix spellings with a lookup record (like your NameFixes)
    //    values not in the record stay as they are
    //    example: Clean[MapValues](T, {"Country"}, [#"Burma" = "Myanmar", #"Swaziland" = "Eswatini"])
    // ---------------------------------------------------------------
    fnMapValues = (tbl as table, columns as list, mapping as record) as table =>
        Table.TransformColumns(tbl, List.Transform(columns, (c) => {c, each if _ is text then Record.FieldOrDefault(mapping, _, _) else _})),

    // ---------------------------------------------------------------
    // 7) StandardizeCase: make category text consistent: "lower" (default), "upper", "proper"
    //    "Paid Search", "PAID SEARCH", "paid search"  ->  "paid search"
    // ---------------------------------------------------------------
    fnStandardizeCase = (tbl as table, columns as list, optional mode as nullable text) as table =>
        let
            m = if mode = null then "lower" else Text.Lower(mode),
            f = if m = "upper" then Text.Upper else if m = "proper" then Text.Proper else Text.Lower
        in
            Table.TransformColumns(tbl, List.Transform(columns, (c) => {c, each if _ is text then f(_) else _})),

    // ---------------------------------------------------------------
    // 8) PercentToNumber: "61%" -> 0.61  (asFraction = true, default)
    //                     "61%" -> 61    (asFraction = false)
    //    values that are already numbers are kept as they are
    // ---------------------------------------------------------------
    fnPercentToNumber = (tbl as table, columns as list, optional asFraction as nullable logical) as table =>
        let
            frac = if asFraction = null then true else asFraction,
            conv = (v) =>
                if v = null then null
                else if v is number then v
                else
                    let
                        t   = Text.Trim(Text.From(v)),
                        has = Text.EndsWith(t, "%"),
                        n   = Number.From(Text.Remove(t, {"%"}), "en-US")
                    in
                        if has and frac then n / 100 else n,
            Converted = Table.TransformColumns(tbl, List.Transform(columns, (c) => {c, conv}))
        in
            Table.TransformColumnTypes(Converted, List.Transform(columns, each {_, type number})),

    // ---------------------------------------------------------------
    // 9) CountDuplicates: how many duplicate rows (for your Cleaning Log)
    //     RemoveDuplicates: remove them
    //     keyColumns = null -> full-row duplicates; or give key columns
    //     ALWAYS count and look first: identical rows are not always errors
    // ---------------------------------------------------------------
    fnCountDuplicates = (tbl as table, optional keyColumns as nullable list) as number =>
        Table.RowCount(tbl) - Table.RowCount(if keyColumns = null then Table.Distinct(tbl) else Table.Distinct(tbl, keyColumns)),

    fnRemoveDuplicates = (tbl as table, optional keyColumns as nullable list) as table =>
        if keyColumns = null then Table.Distinct(tbl) else Table.Distinct(tbl, keyColumns),

    // ---------------------------------------------------------------
    // 10) AddBucket: category column + order column (e.g. GDP or population bands)
    //     upperBounds: {1e10, 1e11, 1e12}
    //     labels     : {"Low", "Medium", "High", "Very High"}   (one more label than bounds)
    //     value < 1e10 -> "Low", < 1e11 -> "Medium", < 1e12 -> "High", else "Very High"
    //     null -> "No Data"; also adds "<newColumn> Order" for sorting in Power BI
    // ---------------------------------------------------------------
    fnAddBucket = (tbl as table, sourceColumn as text, newColumn as text, upperBounds as list, labels as list, optional noDataLabel as nullable text) as table =>
        let
            nd    = if noDataLabel = null then "No Data" else noDataLabel,
            pick  = (v) =>
                if v = null then nd
                else
                    let pos = List.PositionOf(List.Transform(upperBounds, (b) => v < b), true)
                    in  if pos = -1 then List.Last(labels) else labels{pos},
            order = List.Combine({labels, {nd}}),
            Added = Table.AddColumn(tbl, newColumn, each pick(Record.Field(_, sourceColumn)), type text)
        in
            Table.AddColumn(Added, newColumn & " Order", each List.PositionOf(order, Record.Field(_, newColumn)) + 1, Int64.Type),

    // ---------------------------------------------------------------
    // 11) WideToLong: one column per brand/year -> one row per item (Unpivot)
    //     NOTE: Table.UnpivotOtherColumns DROPS null cells
    // ---------------------------------------------------------------
    fnWideToLong = (tbl as table, idColumns as list, attributeName as text, valueName as text) as table =>
        Table.UnpivotOtherColumns(tbl, idColumns, attributeName, valueName),

    // ---------------------------------------------------------------
    // 12) Standard: the safe default pipeline for ANY new file
    //     TidyHeaders -> CleanValues -> remove empty rows -> AutoTypes
    //     It does NOT remove duplicates on purpose (decide that yourself)
    // ---------------------------------------------------------------
    fnStandard = (tbl as table, optional culture as nullable text, optional keepAsText as nullable list, optional nullTokens as nullable list) as table =>
        let
            s1 = fnTidyHeaders(tbl),
            s2 = fnCleanValues(s1, nullTokens),
            s3 = Table.SelectRows(s2, each List.NonNullCount(Record.FieldValues(_)) > 0),
            s4 = fnAutoTypes(s3, culture, keepAsText)
        in
            s4,

    Library = [
        LoadCsv          = fnLoadCsv,
        TidyHeaders      = fnTidyHeaders,
        CleanValues      = fnCleanValues,
        AutoTypes        = fnAutoTypes,
        ParseDates       = fnParseDates,
        MapValues        = fnMapValues,
        StandardizeCase  = fnStandardizeCase,
        PercentToNumber  = fnPercentToNumber,
        CountDuplicates  = fnCountDuplicates,
        RemoveDuplicates = fnRemoveDuplicates,
        AddBucket        = fnAddBucket,
        WideToLong       = fnWideToLong,
        Standard         = fnStandard
    ]
in
    Library


// =====================================================================
//  EXAMPLE 1 - Egypt Delivery Trends (weekly file)   [new Blank Query]
// =====================================================================
// let
//     Source  = Clean[LoadCsv]("C:\Data\interest_over_time_weekly.csv", ",", 2),
//     Cleaned = Clean[Standard](Source),
//     Renamed = Table.RenameColumns(Cleaned, {
//                   {"talabat: (Egypt)", "talabat"}, {"breadfast: (Egypt)", "breadfast"},
//                   {"instashop: (Egypt)", "instashop"}, {"elmenus: (Egypt)", "elmenus"}}),
//     Dated   = Clean[ParseDates](Renamed, {"Week"}, {"yyyy-MM-dd"}),
//     Long    = Clean[WideToLong](Dated, {"Week"}, "Brand", "Interest")
// in
//     Long
//
// =====================================================================
//  EXAMPLE 2 - Egypt Delivery Trends (governorate file)
// =====================================================================
// let
//     Source  = Clean[LoadCsv]("C:\Data\governorate_share.csv", ",", 2),
//     Cleaned = Clean[Standard](Source),
//     Short   = Table.TransformColumnNames(Cleaned, each Text.BeforeDelimiter(_, ":")),
//     Pct     = Clean[PercentToNumber](Short, {"talabat", "breadfast", "instashop", "elmenus"}),
//     Long    = Clean[WideToLong](Pct, {"Region"}, "Brand", "Share")
// in
//     Long
//
// =====================================================================
//  EXAMPLE 3 - country data (GDP buckets)
// =====================================================================
// let
//     Source   = Clean[LoadCsv]("C:\Data\Country Data Comparison.csv", ";"),
//     Std      = Clean[Standard](Table.Skip(Source, 1)),
//     NoEU     = Table.SelectRows(Std, each [Country] <> "#European Union"),
//     Names    = Clean[MapValues](NoEU, {"Country"}, [#"Burma" = "Myanmar", #"Swaziland" = "Eswatini"]),
//     GDPCat   = Clean[AddBucket](Names, "GDP", "GDP Category", {1e10, 1e11, 1e12}, {"Low", "Medium", "High", "Very High"})
// in
//     GDPCat
