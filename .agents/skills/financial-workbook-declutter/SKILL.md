---
name: financial-workbook-declutter
description: "Review and simplify complex Excel-based financial analysis workbooks. Remove unnecessary visual clutter, repetitive AI-style explanations, oversized sections and excessive formatting while preserving formulas, financial information, source evidence, accuracy and professional academic presentation."
---

# Financial workbook declutter

Refine an existing finance assignment so its financial tables are easy to read and its explanations are concise. Preserve a successful existing design. Apply the user's selected measures, reporting boundary, units and conventions; do not import choices from an earlier assignment.

## Inspect before editing

1. Locate the latest actual submission and its audit/source records. Inspect every sheet, hidden inputs, formulas, names, links, notes, chart series, print areas and calculation settings. Render representative views of every sheet before selecting changes.
2. Capture a baseline with `scripts/validate_workbook.py snapshot input.xlsx baseline.json`. Retain the original file or a backup. Record expected measure counts, company-year observations, averages, mandatory measures and native charts from the actual workbook rather than assuming counts.
3. Identify dependencies before moving or removing anything. Useful supporting calculations are financial information even if the labels look repetitive.
4. Prioritize actual clutter: competing fonts and colors, large banners, excessive bolding, tall rows, repeated explanations, crowded notes, empty space and internal audit identifiers. Do not remove useful evidence or uncertainty disclosures.

## Edit minimally

Use one font family, modest type, restrained existing colors, simple section headings, light structural borders, readable widths and fitted row heights. Preserve sheet names, hidden openings, formulas, financial formats, chart references and financial table positions unless the user explicitly authorizes structural change.

Shorten useful visible notes. Move repeated detail into the existing Overview or methodology area, preserving its meaning and source location. Keep full source and reconciliation evidence in the audit record. Add a sheet only when the existing structure cannot serve the reader.

Read [references/writing-style.md](references/writing-style.md) when revising methodology. Write what was done, why and any specific limitation in ordinary academic English. Never invent a business cause or imply an unknown published formula has been confirmed.

Verified inputs, prices, dividends, corporate actions, balance conventions, signed accounting inputs, `n/m` rules and sample averages are protected. Genuine zeros stay numeric and remain in applicable averages. A label saying no dividend was paid must be qualified if cash payments actually settle earlier-year liabilities. Never replace a formula with a number or text for presentation.

An unrelated financial defect discovered during cleanup is not permission to rebuild the model. Preserve it, document its scope and effect, and distinguish presentation completion from financial submission readiness. Correct it only within the user's authorized scope, with official evidence and consistent handling of affected observations.

## Verify and deliver

1. Recalculate in the authoring engine and, where available, the native spreadsheet engine. Preserve intentional chart-gap errors separately from unexpected errors. Do not claim saved caches alone prove recalculation.
2. Capture the revised snapshot and compare it with the baseline. The validator protects formulas, numeric inputs, cached numerical results, sheet order/visibility, hidden inputs, names, chart parts and source annotations. Supply an explicit reviewed text-change allowlist; investigate every other difference.
3. Independently check required ratios and averages, including numeric zero and `n/m` handling. Confirm price/share bases and all original chart references remain intact. Test any authorized calculation change separately.
4. Render every changed sheet, inspect normal-size text and chart readability, and check print areas and page breaks. Do not shrink type to conceal crowding.
5. Deliver the revised workbook and a short account of actual changes, validation coverage and remaining issues. State whether any financial formula or result changed. Keep internal audit files unchanged unless separately authorized.

### Completion checklist

- [ ] Latest input and backup identified; full workbook inspected.
- [ ] Selected and mandatory measures, financial statements and hidden openings retained.
- [ ] All formulas, numerical inputs and financial results preserved or explicitly reconciled.
- [ ] Zero, unavailable and non-meaningful observations remain distinct.
- [ ] Sample averages and native charts verified against the baseline.
- [ ] Source relationships, notes, reconciliation evidence and material limitations retained.
- [ ] Recalculation and rendered layout checks completed, with limitations disclosed.
- [ ] No unqualified submission-ready claim while a known financial error remains.
