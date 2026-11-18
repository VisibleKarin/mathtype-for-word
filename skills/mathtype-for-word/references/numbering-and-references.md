# MathType-native numbering and references

## Contents

1. Required default
2. Native number anatomy
3. Native reference anatomy
4. Why Word numbering is not equivalent
5. Editing and field updates
6. Default scope

## Legacy renderer default

The bundled `config/defaults.json` expresses the requested Format Equation Numbers state:

| Setting | Value |
|---|---|
| Mode | Simple Format |
| Chapter Number | Off |
| Section Number | Off |
| Equation Number | On, Arabic |
| Enclosure | Parentheses |
| Separator | Off |
| Change format for | New equation numbers and Whole document |
| Update automatically | On |
| Warn on first equation number | On |
| Warn on references | Off |
| Default for new documents | On in this agent workflow |

The general marker renderer normalizes the visible number to `(1)`, `(2)`, and so on, even if the interactive MathType dialog previously used a section-aware form such as `(1.1)`. This is a renderer policy, not a limitation of native MathType numbering. Package-local edits preserve the original native fields and scheme.

## Native number anatomy

MathType 7 uses nested Word fields. The important visible-number structure is conceptually:

```text
MACROBUTTON MTPlaceRef
  SEQ MTEqn \h
  (
  SEQ MTEqn \c \* Arabic
  )
```

The hidden sequence field advances the MathType equation counter. The current-value field displays it. The outer `MTPlaceRef` macrobutton is the target a user clicks after inserting a reference placeholder.

MathType may also insert a document setup macrobutton containing hidden reset fields. Preserve it. Do not flatten or unlink these fields.

The bridge calls `MTCommand_InsertEqnNum` first, so MathType creates its full native structure. It then removes only current chapter or section components and their separator from the visible `MTPlaceRef` field. It does not replace the native mechanism.

## Native reference anatomy

The interactive and automated workflows are the same:

1. `MTCommand_InsertEqnRef` inserts `equation reference goes here` and creates the temporary `MTReference` bookmark.
2. The target equation number's `MTPlaceRef` action runs.
3. MathType creates a target bookmark named like `ZEqnNum394416`.
4. The placeholder becomes a `GOTOBUTTON` field containing a nested `REF` field.

Conceptual field:

```text
GOTOBUTTON ZEqnNum394416
  REF ZEqnNum394416 \* Charformat \! \* MERGEFORMAT
```

This is why a literal `(1)` or a Word list cross-reference is not acceptable. It lacks MathType's target placement behavior and bookmark relationship.

## Why Word numbering is not equivalent

Do not use these for the primary workflow:

- Numbered lists.
- Captions labeled Equation.
- Manually authored `SEQ Equation` fields.
- Manually typed numbers.
- `\label` and `\ref` text imported as TeX.

They may look similar but are not MathType's numbering/reference implementation. Use them only after a probe failure and explicit user acceptance of a fallback.

## Editing and field updates

After adding, deleting, or moving numbered equations:

1. Update all Word fields.
2. Verify the `MTEqn` current values are sequential from 1.
3. Verify every `GOTOBUTTON` target bookmark exists.
4. Search for `Error! Reference source not found.`
5. Validate again with the original manifest or an updated manifest.

Never edit `ZEqnNum...` names manually. MathType owns those bookmarks.

## Default scope

`configure_mathtype_word_defaults` writes the packaged default profile to `%APPDATA%\MathTypeForWordAgent\defaults.json` and sets MathType's user warning preferences. Every document processed by this bridge uses the profile and is normalized to the requested number form.

The skill does not automate the localized Format Equation Numbers dialog with keystrokes. Documents created manually outside this workflow retain the user's interactive MathType configuration unless the user also selects the same settings in the dialog.

## Verified chapter and section test

MathType-native equation numbers are themselves implemented with Word fields. A field-based result does not imply ordinary Word list/caption numbering. In this project, `MTCommand_InsertEqnNum`, `MTCommand_InsertEqnRef` and the target's `MTPlaceRef` action generate the native fields/bookmarks. See the [official Word tips](https://docs.wiris.com/en_US/mathtype-tips-tricks-category/tips-to-use-with-microsoft-word).

An opt-in synthetic test now exercises the installed MathType implementation without `Normalize-SimpleNumberField`:

```powershell
python tests/run-native-numbering-integration.py --output-dir <new-absolute-folder>
```

The test uses MathType's `MTLib.WriteDocPropString` to save the document-local `MTEquationNumber2` preference. MathType reads it and constructs the number components itself; no global numbering defaults are changed. Tested formats:

| Document-local format | Saved visible numbers |
|---|---|
| `(#E1)` | `(1)` through `(6)` |
| `(#C1-#E1)` | `(1-1)`, `(1-2)`, `(1-3)`, `(2-1)`, `(2-2)`, `(2-3)` |
| `(#C1.#S1.#E1)` | `(1.1.1)`, `(1.1.2)`, `(1.2.1)`, `(2.1.1)`, `(2.1.2)`, `(2.2.1)` |

`MTCommand_InsertNextChapter` increments the equation chapter, resets section to 1 and restarts equations. `MTCommand_InsertNextSection` increments the equation section and restarts equations. The native break is an `MTEditEquationSection2` macrobutton holding hidden `MTChap`, `MTSec` and/or `MTEqn` sequence fields. It is independent of Word's page/section breaks: all test documents have exactly one Word section. On the MathType ribbon, these correspond to Chapters & Sections → Insert Next Chapter / Insert Next Section.

Each output has six editable `Equation.DSMT4` objects, six native numbers, six native reference occurrences, no OMath and no Word list-numbered paragraphs. A fourth output swaps the first two equation paragraphs; the existing references follow their bookmarks, so the reference to the addition equation changes from `(1.1.1)` to `(1.1.2)`. Saved native data must equal the original equations in the new order. The native update command is `MTCommand_UpdateEqns` on the tested installation.

To see the hidden equation chapter/section labels in Word, use Show/Hide ¶. Place new equation breaks outside heading paragraphs used for a table of contents, because copying such headings into a TOC can duplicate breaks; see [official chapter/section guidance](https://docs.wiris.com/en_US/editing-formatting-equations/equation-chapter-section-numbers-suddenly-change).

These results are live tested on Word 16.0 / MathType 7.8.2.441. They demonstrate a native route in the reproducible test; the general render/validate/update MCP interfaces have not yet gained a chapter/section manifest contract. Do not run their simple-number normalization/validator on these chapter-aware outputs, or claim the generic MCP interface already supports this mode. The More Breaks / Modify Break dialogs and arbitrary explicit starting chapter/section values have not been automated by this test.
