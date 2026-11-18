# Word conversion and local layout policy

This fork follows the user's manuscript preferences. These are workflow choices, not universal Word typesetting rules. Scope every change to the requested equations and nearby text. Do not implement the future formula-writing adviser or illustrations as part of a conversion task.

## OMML and existing documents

Use a working copy; a separate output filename protects the original and does not mean rebuilding the manuscript. Scan first. `scan_mathtype_candidates` returns a source SHA-256 and `word_equation_objects` with one-based `omml_index`, paragraph and readable text. Inspect the source OMML XML where fractions, matrices or scripts are ambiguous; linear text is not a complete structural representation. Review TeX per selected formula before conversion.

`mark_mathtype_document` accepts reviewed OMML targets as well as existing plain-text marks:

```json
{"schema_version":1,"source_sha256":"<scan hash>","marks":[
  {"omml_index":1,"marker":"{{MATH:energy}}"}
]}
```

OMML targets require the unchanged source hash. The bridge replaces the complete OMML contents before removing the container, preventing leftover superscripts/subscripts when Word would otherwise collapse the original range. Only selected objects become markers; all markers must be unique. Fields/comments inside a target are refused for manual range review. Then render the reviewed TeX manifest. Normal surrounding Word text is not a conversion target. Strict render validation requires no remaining OMath by default; a local conversion must explicitly declare the unselected count with `expected_remaining_omml`. Account for any OMML intentionally simplified into ordinary characters in a separate reviewed edit.

For an existing mixed paragraph (equation, number, soft break, explanation), keep that paragraph. If its native number already exists, use `numbering: preserve` and insert only the new object using `layout: inline, numbered: false`, preserving the existing field and tab structure, or use the local donor replacement workflow for existing MathType objects. Here `inline` controls insertion without paragraph restructuring; it does not reclassify a visibly displayed formula as prose. Do not add another number at the physical paragraph end after the explanation. A new display marker must be alone in its paragraph. Unsupported ranges should be reported, not silently split.

## Size

Unless preservation is requested, match Full to actual destination body characters. Use numeric pt from an explicit template/user choice, or `full_size_pt: "match_context"` in the local formatting plan. The worker resolves all contextual sizes before importing any carrier object/styles and returns source paragraph, style, chosen pt and selection basis. The nearby plain paragraph is inspected character by character, so direct formatting takes precedence over Normal's configured size. Headings, fields, OLE/OMML and script characters are excluded; ambiguity requires explicit pt. Check whether the chosen paragraph is appropriate for the document.

For donor replacement, choose size from the destination, not a temporary donor's default Normal style. Preserve the user's request to retain source size by skipping the size-format operation. Do not additionally shrink numerator/denominator fonts to imitate OMML; preserve MathType's own fraction layout and the existing non-Full preferences.

## Inline characters and baseline

Prefer plain Word text with true Word superscript/subscript attributes for simple inline variables or short expressions. Preserve punctuation, italics/upright roles and script scope. Complex fractions/matrices may need an editable MathType inline object.

Inspect these independently before repairing height/alignment: run baseline (`w:position`), superscript/subscript (`w:vertAlign`), paragraph line rule/height, paragraph text alignment, paragraph grid snapping, character grid snapping, and section document-grid type/pitch. A section's grid definition is separate from either snapping setting. The supplied user examples demonstrate that changing the grid definition can change spacing even with paragraph snapping disabled; they do not prove one universal setting is correct.

Never clear all run baselines, reapply a whole paragraph style, change Normal or disable the section grid globally as a routine formula repair. For a replaced MathType object, import its matching preview, dimensions and new baseline together. Keep ordinary neighboring characters unchanged. Correct a plain-text variable's accidental baseline only in the selected run and according to its context; never replace genuine superscript/subscript structure with a visual offset. Test the relevant line/page after changes.

## Display, tabs and continuation

Use genuine MathType display objects. For a new independent display, the bridge computes usable width in points from the current page/column/cell and paragraph indents, and places a center tab at half the width and a right tab at the end. This avoids a character-unit calculation tied to Normal's size. Unequal-width columns need manual layout review. Keep the MathType-provided object baseline; a font-size/position reset over the whole paragraph can erase it.

When the explanation semantically continues the same natural paragraph, ask whether the user wants a soft break (Shift+Enter). Do not submit an assumed choice. Upon confirmation, retain one paragraph with flush-left continuation (no hanging/left indent); a first-line indent acts only on the paragraph's first line. Changing a hard paragraph break into a soft break affects paragraph spacing and style scope, so preserve neighboring anchors/fields and review that exact line. When the user does not authorize joining/splitting, keep existing soft/hard breaks. The local replacement tools preserve existing mixed layouts rather than importing donor paragraph styles.

Only relax exact line spacing or exact table-row height when the equation would clip and the task allows it. The legacy repair disables paragraph grid snapping locally when needed for its at-least line-height fix; preserve section grid definitions and unrelated paragraphs. If exact height is a hard requirement, report the conflict for review. Center vertical text alignment / standard font position may help a particular MathType display, as in the supplied examples, but are not universal inline defaults.
