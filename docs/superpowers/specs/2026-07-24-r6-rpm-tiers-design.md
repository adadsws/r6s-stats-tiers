# R6 RPM Tier Rows And In-Card Video Tier Design

## Goal

Update the automatic-firearm RPM workbook to use the same traditional
horizontal tier-list structure as `r6_operator_tier_chart.xlsx`: RPM-derived
Roman-numeral tiers appear as merged bands on the left, while each operator
card shows its Athieno video tier in a compact badge beside the operator name.
Keep every `补丁说明` worksheet unfrozen.

## RPM Tier Rules

For each operator, calculate the maximum RPM across all primary and secondary
automatic weapons already parsed into the source workbook.

Use these fixed thresholds:

| Tier | Maximum automatic RPM |
| --- | --- |
| `Ⅰ` | `>= 1000` |
| `Ⅱ` | `850-999` |
| `Ⅲ` | `750-849` |
| `Ⅳ` | `< 750`, or no automatic weapon |

These thresholds are fixed constants. They were selected from the current 77
operators but must not be recalculated when source data changes.

Current distribution:

| Tier | Operators |
| --- | ---: |
| `Ⅰ` | 19 |
| `Ⅱ` | 19 |
| `Ⅲ` | 18 |
| `Ⅳ` | 21 |

## Sorting And Display

- Keep separate attacker and defender RPM worksheets.
- Group operators into left-side RPM tier bands in this order:
  `Ⅰ`, `Ⅱ`, `Ⅲ`, `Ⅳ`.
- Use red for `Ⅰ`, orange for `Ⅱ`, yellow for `Ⅲ`, and gray for `Ⅳ`.
- Within each RPM tier, sort operators by maximum automatic RPM descending.
- Preserve source order when maximum RPM values are equal.
- Show no more than five operator cards per row. Wrap additional cards onto
  another row inside the same RPM tier, matching the traditional Tier List.
- Operators without automatic weapons remain after all operators with a
  numeric RPM and receive tier `Ⅳ`.
- Render the operator's Athieno video tier `S/A/B/C/D/F` in a compact colored
  badge immediately to the right of the operator name. Use the existing video
  tier colors from `r6_tiers.py`.
- The in-card badge must not show the Roman RPM tier.
- Do not add a dedicated full-height tier cell to an operator card.
- Do not display the derived maximum RPM or a numeric rank separately; existing
  primary and secondary RPM facts remain on each card.
- Do not change Athieno scores, the statistics workbook score cells, or the
  traditional Tier List workbook.

## Patch Worksheet

`r6_patch_notes.add_patch_notes_sheet()` must leave `freeze_panes` unset on the
`补丁说明` worksheet. The title, source links, patch table, colors, filtering,
printing, and data-status text remain unchanged.

## Code Boundaries

- `r6_rpm_chart.py` owns RPM thresholds, classification, grouping, sorting,
  left-side Roman tier bands, compact in-card video-tier badges, and
  validation.
- `r6_patch_notes.py` owns the patch worksheet freeze-pane behavior.
- Existing operator and weapon parsing remain unchanged.
- The bundled project Skill copies of modified production scripts must remain
  byte-for-byte synchronized with project scripts.

## Error Handling

- Reject nonnumeric maximum RPM values before classification.
- Every operator must resolve to exactly one RPM tier in
  `Ⅰ`, `Ⅱ`, `Ⅲ`, `Ⅳ`.
- Every rendered RPM card must carry a valid Athieno video tier in
  `S`, `A`, `B`, `C`, `D`, `F`.
- Operators without an automatic weapon are valid and classify as `Ⅳ`.
- Existing workbook, icon, score, and source-data validation remains active.

## Tests And Verification

- Unit-test every threshold edge: `1000`, `999`, `850`, `849`, `750`, `749`,
  and no automatic weapon.
- Verify RPM sorting still uses the maximum across primary and secondary
  automatic weapons and preserves source order for ties.
- Verify the left-side merged tier labels are exactly `Ⅰ`, `Ⅱ`, `Ⅲ`, `Ⅳ` from
  top to bottom and use the approved RPM colors.
- Verify each RPM tier wraps after five operators and retains all operators.
- Verify every RPM card contains one compact `S/A/B/C/D/F` video-tier badge
  immediately after the operator name and uses the existing Athieno colors.
- Verify no operator card contains a Roman-numeral badge or a full-height tier
  cell.
- Verify `补丁说明.freeze_panes` is unset in all generated workbooks.
- Run the complete `unittest` suite and Skill validator.
- Regenerate all three workbooks because the shared patch worksheet changes.
- Audit workbook sheet order, status rows, links, image counts, and RPM tier
  distribution.
- Open all three final workbooks read-only in Microsoft Excel and visually
  inspect the RPM chart and patch worksheet.
