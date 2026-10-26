# Excel data transfer skill

A configurable desktop Excel VBA implementation. The files contain complete procedures, not pseudocode. No Excel automation runtime is installed in the authoring workspace: the VBA has not been compiled or executed here. Run **Debug > Compile VBAProject**, then `RunTransferTests` in your supported Excel installation before using real workbooks.

## Setup

1. Create a dedicated `.xlsm` controller outside OLD and NEW. Open the VBA editor (Alt+F11), File > Import File: import `vba/Transfer.bas`. No added VBA references are needed. Enable macros only for your trusted controller.
2. Run `SetupTransfer` once. It creates/replaces **Transfer Config in the controller**. Do not run setup again after configuring it unless you want to reset configuration.
3. Run `SelectWorkbooks`, or enter full local paths in B2 (OLD), B3 (NEW template), B4 (new output). The output must not exist and must have the template's extension. Close OLD and NEW first. Use local synced files rather than URL paths; open passwords, sensitivity-label dialogs and Protected View require preparation in Excel.
4. Set B5=True for a saved dry run; B6=False for links (True for values); B7=False for case-insensitive keys. B8 optionally specifies a fresh external audit `.xlsx` path. Output is a file copy, never a conversion of the original. Original files are never saved by the tool.
5. Replace the example rows from row 11. Enter **both** sides for each logical region. Rows must be contiguous through the last populated A cell. Booleans must be Excel TRUE/FALSE values.

| Column | Meaning |
|---|---|
| A | OLD or NEW |
| B | Exact worksheet name on this side; paired configuration rows provide explicit sheet mapping |
| C | Globally unique stable region ID per side; same ID on corresponding OLD/NEW rows |
| D | Single rectangular A1 range, worksheet/workbook named range, or Excel Table name |
| E | Number of column-header rows at top of region, EXPLICIT mode |
| F | Number of row-header columns at left of region, EXPLICIT mode |
| G | Designated input address; may be a union of areas within the body, never headers |
| H | Optional section description for configuration documentation |
| I | Optional Excel RGB Long integer, e.g. obtained from `RGB(r,g,b)`; no implicit default colour |
| J | TRUE: scan `DisplayFormat.Interior.Color`; FALSE: scan `Interior.Color` |
| K | Include hidden inputs/sheet; default TRUE. FALSE excludes hidden rows, columns and sheets from that region's input inventory |
| L | Include worksheet. A FALSE row excludes the whole worksheet on that side |
| M | EXPLICIT or AUTO |
| N | Human note |

Use EXPLICIT for merged/multi-level headers. E=2, F=2 means two header rows and two identifier columns. Header labels are read from merged top-left cells and expanded **in memory**. Never merge, unmerge or fill worksheet cells. Merged body/input cells are rejected. Multi-cell spill/array destinations are not supported; the write/rollback path handles an Excel rejection without replacing the spill owner.

P:R holds approved **component-label aliases**: axis R/C, raw label, approved replacement. Normalisation replaces NBSP and CR/LF with spaces, trims outer spaces and optionally lowercases; it preserves punctuation, internal spaces, units, and text leading zeros. Keys are length-prefixed components (not delimiter concatenation), so punctuation cannot create component collisions. Numeric header values use Value2 serials; store date and zero-padded identifiers as explicit text when displayed formatting has semantic meaning. Duplicate composite keys, including normalisation/alias collisions, invalidate the whole region. Alias replacement is one-level, global per axis, not recursive.

## Discovery and matching

All included worksheets on each side are inspected. Explicit configuration comes first, then Tables and range-valued Names. UsedRange only bounds a content-based Find search; formatting-only tails do not become discovered regions. Unprotected worksheets use CurrentRegion at content seeds with address deduplication. Protected worksheets never call CurrentRegion; a read-only fallback scans blank row/column boundaries. This fallback is intentionally conservative and can group adjacent tables or split tables with internal blank rows.

Automatically found blocks are **candidate regions only**. They are inventoried and logged for configuration with unknown coverage; no automatic block is a writable designation. Candidate counts and validated counts are separate. Titles, notes, adjacent tables, blank gaps, merged headings and formula-empty cells can defeat inference. This tool does not guarantee discovery of every logical table. A blank input omitted from all explicit designations cannot be discovered reliably. Completeness means the reviewed configuration covers the intended domain, not every possible input in a workbook.

For configured AUTO regions the first row/column undergo strict nonblank, nonformula text checks, followed by full composite uniqueness checks. If positional detection fails, the optional colour scanner requires exactly one continuous labelled header row and column; preceding title rows/columns are cropped in memory. Colour never establishes identity by itself. AUTO does not infer multi-level hierarchies: configure those explicitly. Semantically plausible but incorrect headers still need human preflight review.

Region IDs in the paired rows are the authoritative explicit mapping, including renamed worksheets, renamed tables, section titles and moved regions. Set the IDs from stable table/name/section identities after review. Section descriptions do **not** trigger fuzzy automatic matching. Header-signature-only and section-title-only pairing are deliberately not implemented: require an explicit ID for those cases. Missing or multiple destinations are logged, never paired by sequence.

Each designated source input supplies composite row/column keys. Destination Find searches have explicit LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte and SearchFormat; literal wildcards are escaped. All hits are enumerated with wraparound detection, merged spans supply candidate indexes, complete composite keys validate candidates, and only one row/column answer is accepted. Resolutions are cached. Cells are resolved through `destinationWorksheet.Cells(row, column)` and checked against the designated body/input ranges. Conflicting source claims block the earlier proposal as well. Reconciliation inventories NEW independently so absent sources are visible.

## Execution and review

A run always builds all proposals before writes. A B5=True dry run saves inventories, proposals, exceptions and summaries without writing input contents. For execution choose a **new output filename**, set B5=False, and run again. Only cells in designated input areas that are unlocked and have no unrelated formula can change. Nonformula input constants may be replaced intentionally. Existing business formulas, including unrelated external links, are rejected.

Events and table formula autofill are disabled during controlled writes. Calculation is manual, source links open with UpdateLinks=0, source/template macros are disabled on open, and only written cells are explicitly calculated. Events, calculation, screen updating, table autofill and macro security are restored. No clipboard, Select, Selection or ActiveSheet is used. No protection API is called in the production module.

Link mode uses an absolute escaped external reference wrapped in `IF(ISBLANK(ref),"",ref)`. Physical source blanks display an empty string rather than zero. **A formula returning an empty string is not physically empty.** Genuine zeros, text including leading zeros, Boolean values and date serials retain their Value2 semantics. Values-only clears physical blanks and writes strings as literal text, including strings beginning `=`. Source formula results are their current cached/calculated values; the tool does not refresh dependent queries or source links. Source errors are logged and skipped. Different 1900/1904 date systems are rejected, rather than silently shifting dates.

Validation support: no-validation cells, input-only validation, literal lists and constant numeric bounds (whole, decimal, date, time, text length). Custom formulas, formula/range lists and formula bounds are **rejected as invalid value**, because safe prospective evaluation in the destination context is not implemented. This remains true even for warning-only rules. Add a reviewed validation adapter if needed; do not disable validation. Date/time bounds must be numeric serials.

Audit sheets are the only structural additions: **missing source to link**, **transfer audit**, **region inventory**, **transfer plan**, **transfer summary**. They are excluded from discovery. Existing compatible logs append run IDs/timestamps and are never cleared. Header schemas are checked. Protected audit sheets are conservatively considered nonwritable; structure protection or blocked internal sheets causes an external audit workbook, whose full path is recorded in the summary. Protecting structure is never bypassed. External audit creation failure is fatal before input processing.

The exception sheet includes source/destination identifiers, raw/key labels when reliable, detection method, stage/category, error details, unchanged/restored outcome, coverage and corrective action. Invalid configured regions log each identifiable input as unresolved; missing worksheets/unknown boundaries explicitly carry unknown coverage. Corrections usually mean narrowing/configuring boundaries, disambiguating labels, adding a missing counterpart, or fixing eligibility through the workbook owner's normal process.

Source and destination intended-input/unresolved counts are separate. Unknown coverage prevents the verified status. Dry runs do not count proposals as successful transfers. Completed status is saved in the summary; external audit path is shown there. A fatal run preserves available audit information and leaves the partial output open for inspection; it is never announced as complete. Logging failure causes an explicit error dialog and stops processing. Do not use a partially saved output as a trusted template.

## Verification, preservation and reruns

Each successful write checks exact formula text and typed Value2 semantics. Failure restores prior content where possible, logs restoration and continues; restoration failure stops. Production code never writes to input-cell formatting, validation or protection properties. A preflight preservation snapshot checks nonpermitted formulas/constants, lock/hidden settings, merged addresses, standard fills/fonts/alignment/borders/number formats/styles, row/column dimensions, validation properties, protection options and sheet count. It checks after each proposal and again at the end, then reconciles every identified input.

Limit: this property snapshot is not a byte-for-byte OOXML verifier. It checks conditional-formatting **rule count**, not every rule definition, and does not fingerprint rich-text runs, gradient internals, drawings, charts, VBA projects, digital signatures, external add-in behaviour or all newer Excel metadata. Such content is never intentionally edited, but this implementation cannot certify its preservation. Use independent package-level comparison for workbooks requiring that guarantee. Very large snapshots (>1,000,000 UsedRange cells per sheet) stop safely; discovery/Find and preservation checks favour clarity over high-volume speed. Full-sheet property scans can be slow. This implementation is therefore a conservative configured baseline, with these explicit limits, not universal automatic layout recognition or exhaustive workbook-integrity certification.

Rerun on a fresh copy of the **previous successful output**: set that file as NEW template and choose a new output path. Only an internal `transfer audit` row with successful outcome, link mode, exact destination and exact current formula authorises replacing a previous link. Same formula/value becomes “already correct”; altered or unrelated formulas remain blocked. Audit records are trusted local evidence, not cryptographically authenticated. External-only audit ownership is not automatically loaded; retain the audit and manually review before preparing an approved internal audit template. Values-only reruns deliberately rewrite eligible constants. Do not delete old successful audit records to make a formula appear eligible.

## External-link dependencies

The formula embeds OLD's path, workbook name, worksheet and absolute cell address. OLD must remain accessible and trustworthy. Moving or renaming OLD can break the reference or require Excel's Change Source workflow. Changes in OLD may flow through on Excel link refresh/recalculation; the audit documents the original mapping, not a frozen source value. An unavailable OLD may leave cached results, stale values, a security prompt or a link error; caches are not proof of a current successful transfer. Closed-workbook refresh behaviour and trust/security policies vary by Excel version. Do not enable unrelated updates just to clear a warning. Values-only outputs have no newly created external dependency. Rebuild from a stable archived OLD file when reproducibility matters.

## Tests

Import `vba/TransferTests.bas`, save the controller and run `RunTransferTests`. It creates timestamped fixtures under `%TEMP%`, performs dry-run, link, rerun, blocked-audit and values-only runs, and raises assertions on concrete values/logs/preservation. It replaces the controller's config, not user workbooks. It leaves fixtures for review. See `tests/EXPECTED.md`. No tests have been represented as Excel-passed in this package.
