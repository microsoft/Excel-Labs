# Frequently asked questions and problems

#### What is the very-hidden sheet used for?
AFE uses the sheet to detect certain formula locale properties that cannot be extracted from the Office JavaScript API. You can safely delete the sheet at anytime. AFE will recreate the sheet as required when opened. Previous versions of AFE used a hidden sheet to store formulas, this is now deprecated.

#### I see a "The argument is invalid or missing or has an incorrect format" error when trying to save.
This happens when AFE tries to save a formula with certain errors in, such as a missing bracket. To resolve this issue we first recommend the follow steps:
- Check that none of the formulas have errors (red underlines).
- If the formulas use structured references, check that they match the tables and table columns in the workbook.
- If the error persists, feel free to create an issue on the tracker with the formulas.

#### Why does a formula work in the grid but fail in the AFE debugger?
The formula debugger is a **separate, partial evaluator** inside Advanced Formula Environment. It is not the same as Excel’s calculation engine or the classic **Evaluate Formula** dialog.

Excel may return a correct result in the cell while the debugger reports that a feature is not implemented. Common cases include:

- **REGEX functions** (`REGEXEXTRACT`, `REGEXTEST`, `REGEXREPLACE`, `REGEXEXTRACTALL`) — tracked in [Excel-Labs#45](https://github.com/microsoft/Excel-Labs/issues/45).
- **Other unimplemented functions** — the debugger may show a message such as *The formula uses features which the debugger does not yet implement* with a `Details:` function name (see [Excel-Labs#41](https://github.com/microsoft/Excel-Labs/issues/41)).
- **Dynamic / spilled arrays** — the debugger may refuse formulas that use spilled-array features even when the grid result is a single value (see [advanced-formula-environment#83](https://github.com/microsoft/advanced-formula-environment/issues/83)).

**Workaround while debugging:** use Excel’s **Formulas → Evaluate Formula**, or break the calculation into helper cells / intermediate named steps that avoid unsupported debugger features. Grid results remain authoritative when they differ from the debugger.

#### Do module function names allow dots?
Prefer **underscores** in named formula identifiers authored in modules (for example `JSON_GET`), not dotted names such as `JSON.GET`. Dotted identifiers are easy to confuse with module qualification and can break module syntax. Path *arguments* (for example `"meta.captured_at"`) may still contain dots.
