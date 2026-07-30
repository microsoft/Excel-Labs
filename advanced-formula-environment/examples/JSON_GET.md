### Description
Lightweight JSON text helpers for Excel 365, built with `LAMBDA` + `REGEXEXTRACT`. Useful for reading flat property values and simple nested objects from JSON pasted into a range.

**Note:** These formulas calculate in the Excel grid when REGEX functions are available. The AFE **debugger** does not yet implement REGEX functions ([Excel-Labs#45](https://github.com/microsoft/Excel-Labs/issues/45)); use Evaluate Formula or helper cells to step them. Function names use underscores (`JSON_GET`), not dots (`JSON.GET`).

`JSON_OBJECT_AT` only supports **flat** objects (`\{[^}]*\}`). Nested braces inside an object will not match.

### Code
```
// JSON_TEXT — join a multi-cell paste into one string
JSON_TEXT=LAMBDA(json_range,
  LET(
    joined, TEXTJOIN("", TRUE, json_range),
    IF(joined = "", NA(), TRIM(joined))
  )
);

// JSON_PROP — read a scalar property from an object fragment
JSON_PROP=LAMBDA(object_text, prop_name,
  LET(
    pat, """" & prop_name & """\s*:\s*(""[^""\\]*(?:\\.[^""\\]*)*""|true|false|null|-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)",
    raw, IFERROR(REGEXEXTRACT(object_text, pat, 1), NA()),
    IF(ISNA(raw),
      NA(),
      IF(LEFT(raw, 1) = """",
        MID(raw, 2, LEN(raw) - 2),
        IF(raw = "true",
          TRUE,
          IF(raw = "false",
            FALSE,
            IF(raw = "null", NA(), VALUE(raw))
          )
        )
      )
    )
  )
);

// JSON_OBJECT_AT — extract a flat object by key (no nested braces)
JSON_OBJECT_AT=LAMBDA(json_text, key_name,
  LET(
    pat, """" & key_name & """\s*:\s*(\{[^}]*\})",
    IFERROR(REGEXEXTRACT(json_text, pat, 1), NA())
  )
);

// JSON_GET — dotted path into flat nested objects, e.g. "meta.captured_at"
JSON_GET=LAMBDA(json_range_or_text, path,
  LET(
    text, IF(ISREF(json_range_or_text),
      JSON_TEXT(json_range_or_text),
      json_range_or_text
    ),
    parts, TEXTSPLIT(path, "."),
    depth, COLUMNS(parts),
    IF(depth = 1,
      JSON_PROP(text, INDEX(parts, 1)),
      LET(
        obj, REDUCE(
          text,
          DROP(parts, , depth - 1),
          LAMBDA(acc, key, JSON_OBJECT_AT(acc, key))
        ),
        JSON_PROP(obj, INDEX(parts, depth))
      )
    )
  )
);
```

### Sheet usage
```
=JSON_GET(A1, "meta.captured_at")
=JSON_PROP(A2, "elapsed_ms")
```

### Gist link
_(Paste into an AFE module, or publish your own gist and replace this line.)_
