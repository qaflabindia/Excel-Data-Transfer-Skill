---
name: excel-data-transfer
description: Configure and run auditable Excel VBA transfers between workbook versions by worksheet, region and composite header identity while preserving protection and cell metadata.
---

Use this skill when linking designated inputs from an OLD Excel workbook into a separate copy of a NEW template. Read README.md and tests/EXPECTED.md before configuration. Import vba/Transfer.bas into a dedicated trusted .xlsm controller. Import vba/TransferTests.bas to run the fixtures in desktop Excel for Windows.

Workflow:
1. Run SetupTransfer once. Set paths, mode and region/input definitions. Treat the OLD workbook as immutable. Never unprotect either workbook.
2. Give corresponding regions the same globally unique stable ID; the OLD and NEW configuration rows are an explicit worksheet/region mapping. Prefer actual table/named-range names in the range field. Region order and cell addresses are not identities.
3. Configure multi-level headers explicitly. Use AUTO only for simple textual bands or a configured direct/visible header colour. Never approve a discovered candidate simply because it is contiguous or coloured.
4. Run with Dry run=True; inspect region inventory, transfer plan, missing source to link and transfer summary. Resolve ambiguous boundaries and unknown coverage through configuration. Preserve logs.
5. Run with Dry run=False and a fresh output filename. This still performs complete in-memory preflight before any input write. Only valid one-to-one proposals execute; recoverable exceptions continue.
6. Review the final source and destination reconciliations separately. Unknown coverage prevents “Completed and verified”. An external audit is mandatory when internal logging is blocked.
7. Run the fixture harness on the installed Excel version before real use. Report actual test results; do not claim desktop execution based on static inspection.

Safety: never expand writable areas by inference, replace unrelated formulas, disable protection, alter formatting/validation, or update unrelated links. If unsupported validation, uncertain headers, oversized snapshots or unusual workbook behaviour blocks execution, report the limitation and request configuration or a reviewed adapter. Do not weaken eligibility checks to make a run pass.

Ownership: a formula is tool-owned only when the copied template's success audit identifies the exact destination and exact current formula. External-audit ownership is deliberately not imported automatically. Keep external audit evidence for manual review; do not classify arbitrary external links as owned.
