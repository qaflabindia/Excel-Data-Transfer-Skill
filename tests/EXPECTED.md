# Repeatable fixture expectations

Import both modules into a dedicated controller. `RunTransferTests` replaces that controller's config and creates new fixtures, retaining them under a timestamped TEMP directory. It does not require access to the VBA project object model. Run once per second at most (fixture directory names have one-second resolution).

All successful writes also invoke production preservation checks. The fixture assertions check representative preservation properties explicitly. A successful harness run reports its assertion count. A failed assertion shows its description and retains workbooks for investigation. Excel desktop execution and compilation are required; static checks alone do not establish runtime correctness.

| Case | Expected transfer | Expected exception/audit | Preservation checks |
|---|---|---|---|
| Dry run | Input constants remain 99 | Proposed rows exist; no successful-write rows for this run | Business formulas, metadata and protection unchanged |
| Reordered rows/columns | OLD Alpha/Jan B2 resolves to NEW C3; zero to D3; text to B3 | Successful link rows | Headers/labels unchanged |
| Multiple regions | Data F1:I6 maps to H10:K15 via `second` ID | Separate region inventory and success rows | Repeated labels do not cross regions |
| Protected worksheets | Unlocked inputs can link; no unprotect call | B4 locked destination | ProtectContents remains True; locked B4 remains 99 |
| Duplicate labels across regions | Independent valid regions transfer | No cross-region duplicate failure | Input boundaries unchanged |
| Duplicate labels within region | Entire `Duplicate` region skipped | duplicate header; unknown region coverage plus per-input exceptions | Its inputs/headers unchanged |
| Multi-level/merged headers | Hierarchy C3:F5 maps using two-column row keys/two-row column keys | Successful link rows | C1:D1/E1:F1 remain merged, no body merges introduced |
| Colour detection | Title excluded in memory; Colour B3:D5 links | Inventory identifies Interior colour method; title may produce conservative candidate exception | Colour, title and dimensions unchanged |
| Blanks/zeros/dates/text/Boolean | Blank yields `""`, zero stays 0, date stays serial, `00123` stays text, True stays Boolean | Success rows for eligible cells | Existing formats unchanged; values-only blank is physically empty |
| Missing source | DestinationOnly inputs remain 99 | missing source for each | No fabricated source |
| Missing destination | SourceOnly inputs omitted from transfer | missing destination for each | Destination-only rows untouched |
| Missing worksheet/region | Orphan OLD inputs skipped | missing worksheet (unknown coverage); missing region per identifiable orphan input | Other regions proceed |
| Source error | OLD Beta/Mar error is not written | source error | NEW B5 stays original |
| Existing formula | NEW C4 remains `=40+2` | existing formula | Exact formula text retained |
| Validation | NEW D4 rejects value above 1 | invalid value | Same rule and prior content retained |
| Recoverable failure | Controller test hook blocks Data I11; J12 still links | write failure, restored | I11 stays 99; other cells proceed |
| Rerun | Exact audited links are already correct | successful audit outcome `already correct` | C4 still blocked; no unrelated formula ownership |
| Audit sheets blocked | Protected structure unchanged; external audit created | Exceptions/successes in ExternalAudit.xlsx; summary reports its path | No internal audit sheets added, no unprotect |
| Values-only | Literal text preserved, no new formulas, blank clears | Success audit mode values | Same metadata/protection checks |

Additional supported manual regression variants:

1. Set J=True and replace the direct header fill with a conditional-formatting rule that displays the same configured colour; set header Interior fills to no fill. Expected AUTO colour detection through DisplayFormat, same transfers and unchanged rule count. This variant is not asserted by the base harness.
2. Rename OLD to a path containing an apostrophe and use worksheet `O'Brien`. Configure both rows appropriately. Exact link audit must contain doubled apostrophes and calculations must match.
3. Change two header labels to ` A ` and `a` with case sensitivity False. Expect duplicate header for the entire region. Turn case sensitivity True to distinguish them. Use `*`, `?`, `~` in distinct labels to verify literal Find escaping.
4. Make external audit directory unwritable while structure is protected. Expect stopped status before any input writes, error dialog, no verified claim. Restore directory access before rerunning.
5. Delete a required region designation while retaining its blank inputs. Expect candidate/unknown coverage rather than guessed writable cells. Completely blank undesignated areas cannot be discovered: the reviewer must configure them.
6. Use formula-based validation or a different date epoch. Expect invalid value and unchanged destination, never silent conversion.
7. For exhaustive formatting guarantees, independently compare worksheet XML/style/validation/merge/protection definitions in template and output, excluding permitted input formula/value changes and added audit sheets. The built-in property snapshot does not certify every possible OOXML property.
