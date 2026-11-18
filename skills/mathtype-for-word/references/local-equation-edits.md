# Existing Word equations: inspect, format, replace

Use these tools for existing `Equation.DSMT4` objects, especially paragraphs mixing an equation, native number, soft break, explanation and comment anchors. They work on DOCX copies and preserve existing numbering schemes; the legacy marker renderer remains the path for creating new equations and native references.

## Inspect and select

Call `inspect_mathtype_equations(document_path)` first. It reads the DOCX package without Office and returns `source_sha256` plus each equation's `part`, one-based `index` within that part, `native_sha256`, embedding and preview paths. Indices describe this saved input, not permanent object identities. Reinspect after any source edit; Word may also reassign ObjectID/ShapeID and comment IDs during SaveAs. Inspect does not report internal Full size.

Write a local edit plan (different from the marker-render manifest):

```json
{
  "schema_version": 1,
  "source_sha256": "<from source inspection>",
  "operations": [
    {
      "target": {
        "part": "word/document.xml",
        "index": 2,
        "native_sha256": "<from the selected source equation>"
      },
      "full_size_pt": 10.5
    }
  ]
}
```

The size above is an example, not a default. Choose it from the user's/template's body style or explicit request and explain the choice. Avoid guessing from headings, captions, an equation's font size or mixed-size text. Automatic body-style inference is not implemented. `part` defaults to `word/document.xml`. Hashes are required and stale plans fail before publication. Each target occurs once.

## Internal Full size

Call `format_mathtype_equations(input_path, output_path, manifest_path, overwrite=false)`. Each operation contains exactly `target` and `full_size_pt` (finite numeric pt, 1–127). Output must differ from the source; an existing output requires explicit overwrite authorization.

The worker discovers the installed MathType SDK/template, compiles its own helper for the SDK's x86/x64 architecture with the installed .NET Framework compiler, and opens a private hidden Word carrier. It reads the source preferences, changes only Full, transforms MTEF to MTEF without setting a translator, and writes the result to a verified empty native MathType object. Writing into an existing nonempty object is prohibited because native setters can append content while returning success.

After saving and reopening the carrier, it reads internal Full and MathML. Publication requires:

- The readback equals the requested pt and all other saved preferences, including font roles and script ratios, remain unchanged.
- The before/saved MathML trees agree after ignoring size/presentation metadata and annotations. Token order, subscripts/superscripts, fractions, matrix rows/cells and mathvariant remain significant. A same-text structural change or duplicated expression fails.
- Final-package MTEF matches that verified saved carrier, unselected parts remain byte-identical, and XML outside selected objects/baselines is unchanged.

The response includes per-equation `target_full_size_pt`, `saved_full_size_pt`, `mathml_tree_equal`, `other_preferences_preserved`, saved native hashes and changed package parts. Setting Word font size or scaling an OLE object is not equivalent to setting Full. A scale readback of zero is not proof of 100% scaling.

Formatting uses the system clipboard for preference readback, replaces its contents, and shares the Office bridge lock. Do not run it with another Office job. It does not change global defaults, paragraph styles, exact line spacing or row heights. Inspect the output for clipping/collisions and coordinate any needed layout change with the document workflow.

## Local replacement from a donor

Generate and validate a small donor DOCX through the ordinary marker workflow. Set its internal size if required, then inspect both source and donor. Call `replace_mathtype_equations(input_path, output_path, manifest_path, donor_path, overwrite=false)` with:

```json
{
  "schema_version": 1,
  "source_sha256": "<source hash>",
  "donor_sha256": "<donor hash>",
  "operations": [
    {
      "target": {"part": "word/document.xml", "index": 3, "native_sha256": "<old native hash>"},
      "donor": {"part": "word/document.xml", "index": 1, "native_sha256": "<new native hash>"}
    }
  ]
}
```

Replacement is package-only and opens no Office session. It copies only the selected OLE data, matching VML preview, geometry and run baseline. It preserves the source's surrounding text, soft/hard breaks, comments, styles, bookmarks, native number/reference fields and relationship targets. Donor paragraph formatting and chapter-counter resets are not imported. Saved MTEF must match the donor; mathematical correctness and donor typography must be reviewed before this call. This is object replacement, not a claim of general editing within an unchanged OLE identity.

Neither local operation refreshes fields. Existing `(n)`, section/chapter schemes and current values are retained. Use the legacy update/validation tools only if their simple-number contract matches the document; other schemes need an appropriate field-update workflow. Adding/removing/reordering numbers or importing donor references is not implemented by local replacement.

## CLI and limits

From the full plugin root (use absolute paths):

```powershell
python scripts/mathtype_local.py inspect --input $sourceDocx
python scripts/mathtype_local.py format --input $sourceDocx --output $outputDocx --manifest $localPlan
python scripts/mathtype_local.py replace --input $sourceDocx --output $outputDocx --manifest $localPlan --donor $donorDocx
```

- Internal-size formatting currently supports inline equations in the main document, MTEF 5, with the installed native SDK and vendor blank template. Other stories and floating objects require separate work. x86 is live tested; x64 compilation is routed by DLL machine type but is not live tested.
- Package replacement requires a single VML shape/preview and an object-only Word run (other runs in the same paragraph can contain any text/fields/breaks). Shared embeddings/previews, signed packages, differing preview formats and unsupported containers are refused. An object sharing its own run with text/breaks must first be isolated in a working copy; never delete that text.
- The output is published atomically after package validation; failure leaves the source and any existing destination intact. No broad empty-paragraph cleanup is performed: intentional blank paragraphs and anchored breaks are preserved.
- Structure/preservation reports do not prove legibility or mathematical equivalence under arbitrary algebraic rewrites. Inspect a PDF/page render, especially for long fractions, matrices and exact line-height documents. Use the document workflow's existing renders when available.
- Do not publish private manuscripts, exported comments, vendor DLL/WLL/template files or historical logs as fixtures. The repository's local integration uses synthetic text and installed commercial components.
