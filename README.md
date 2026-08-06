# LaTeX Blocks

LaTeX Blocks puts editable LaTeX into Microsoft Office as portable SVG objects. StemTeX performs TeX layout and
SVG rendering; the Office add-in owns placement, editing, and persistence. The source of each block remains on the
visual object itself, so a document continues to display correctly without the renderer installed.

The project contains separate Word and PowerPoint VSTO add-ins. StemTeX is an external infrastructure project, not a
Git submodule of this repository; see [StemTeX integration](docs/STEMTEX_INTEGRATION.md).

## Host capabilities

| Capability | Word | PowerPoint |
| --- | --- | --- |
| Traditional inline formula | Yes — natural width, a standard minimum TeX line box, exact baseline mapping, and a U+2060 boundary scaffold | No |
| Fixed-width LaTeX block | Yes | Yes — positioned slide shape |
| Numbered display equation | Yes — Word-native `SEQ` field and tab stops | No |
| Edit source on the visual object | Yes | Yes |
| Source storage | SVG `AlternativeText` | Shape `AlternativeText` |
| Host font-size integration | Selection-aware inline refresh | Snapshot text size at insertion; explicit block-size control thereafter |
| Text-color integration | Native Word Font Color for inline and numbered formulas; persistent per-Block text color for fixed Blocks | Persistent per-block text-color control |
| Block styling | Fixed Blocks: line spacing, padding, Top/Middle/Bottom vertical placement, text/fill/border color, and border width | Line spacing, padding, Top/Middle/Bottom vertical placement, text/fill/border color, and border width |

Word auto-width content formulas are the only objects surrounded by U+2060 WORD JOINER characters. This keeps a
user-authored ordinary space out of Word's special inline-image spacing path without putting padding or negative
spacing into the TeX SVG. Fixed blocks and numbered equations are not wrapped. The full invariant is specified in
the [Word object model](docs/OBJECT_MODEL.md).

PowerPoint intentionally supports only free-standing blocks. It does not attempt to make an SVG participate in a
text run or simulate inline baseline behavior; see [PowerPoint scope](docs/POWERPOINT_SCOPE.md).

## Use

### Word

The **LaTeX Blocks** Ribbon tab provides **Insert Inline Formula**, **Insert LaTeX Block**, **Insert Numbered
Equation**, **Edit LaTeX Block**, and **Copy as LaTeX**. Select a recognized block to edit its authoritative source.
Select ordinary Word text containing inline LaTeX Blocks and choose **Copy as LaTeX** to place a LaTeX-safe version
on the clipboard: ordinary text is escaped, each recognized inline Block is replaced by its exact author source,
and Word-only U+2060 and numbered-equation scaffolds are omitted. Floating Blocks are not part of Word's selected
text stream and are therefore not included. **Paste from LaTeX** performs the inverse operation for mixed clipboard
text: escaped text characters such as `\%` become ordinary Word text, while real `$...$`, `\(...\)`, `\[...\]`,
`$$...$$`, and standard math environments become editable inline or fixed LaTeX Blocks. It is a mixed-text importer,
not a complete `.tex` document converter. The selected profile is a preference of the Word host, not of an individual
block.

A fixed-width **LaTeX Block** owns an outer frame independently of TeX's layout width. That remains true whether Word
keeps it **In Line with Text** or gives it any floating Layout Option, including **Square**, **In Front of Text**, and
**Behind Text**. After a native resize, LaTeX Blocks rerenders and reframes the SVG instead of preserving Word's
image-scale transform. The commit begins when Word releases the resize gesture; **Reflow Frame** is also available for
an immediate explicit rerender. Moving or rotating a Block does not rerender it. A floating replacement preserves its
position and wrapping, and editing source preserves the outer frame in either representation. Auto-width inline
formulas and numbered equations remain separate, because their Word layout scaffolds are intrinsically inline.

The Word **LaTeX Block** editor has the same block-style controls as the PowerPoint editor: **TeX font size**,
**Line spacing**, uniform **Padding**, **Top / Middle / Bottom** vertical placement, text color, background fill,
and border color/width. These controls apply only to fixed Content Blocks. The authored outer frame minus the
configured padding on all four sides becomes the exact TeX content box. TeX owns zero paragraph indentation,
leading, glyph color, horizontal left alignment, and the selected vertical placement inside that box. Ordinary text
uses a stable TeX line box, so lowercase-only first or final lines retain the font's full ascent/depth; standalone
display math is left untouched. The SVG root places the returned content box at the padding origin and only adds fill,
an inside border, and outer clipping—it does not align the content again. The style is saved with the Block, survives an
InlineShape-to-Shape conversion, and is reapplied after every native frame resize. A new Block initially uses the
insertion point's text color and size; subsequent fixed-Block appearance is edited in the Block editor.
An editor-confirmed default is literal—its displayed 1.20× leading is applied in TeX—while blocks written before the
style editor retain their compatible bare-SVG rendering until they are edited.

**Font Color** in Word's Home tab remains the formula-color control for inline and numbered formulas. To recolor one,
select it and apply Font Color; LaTeX Blocks detects the committed main-button, palette, or More Colors action and
queues one asynchronous SVG refresh immediately. Opening or canceling a picker does nothing. The formula remains an
exact picture selection throughout the visible interaction; the add-in neither polls the document nor changes the
LaTeX source. The same command also works across an ordinary mixed selection such as text–formula–text: Word owns the
text colour, while LaTeX Blocks immediately reconciles every selected formula and preserves the complete range
selection as their SVGs are replaced.

Numbered equations use Word fields. After moving, copying, or deleting complete numbered-equation lines, run
**Update Numbers** (or Word's own field-update command). Deleting only the SVG is not a semantic equation deletion:
the Word field and tab scaffold remain.

### PowerPoint

Use **Insert Block** or **Edit Block**. A new block starts with the ordinary text size at the active caret/selection
when one exists. **Typesetting width (pt)** is an explicit TeX layout control; **TeX size (pt)** rerenders at a
different design size. Every native PowerPoint *size* change also queues a real TeX layout pass: a changed width
derives a new typesetting measure, while a height-only change rerenders at the current measure under the new frame.
Moving or rotating the shape does not render. SVG artwork is never visually scaled. The final SVG viewport always
matches the native PowerPoint frame exactly: after a reflow attempt, content that still exceeds a user-shrunk frame
is clipped rather than enlarging that frame. Use **TeX size (pt)** when the formula itself should become larger or
smaller.

The block editor also exposes **Line spacing**, uniform **Padding**, **Top / Middle / Bottom** vertical placement,
text color, background fill, and border color/width. StemTeX receives the outer frame minus the configured padding
on all four sides as its exact content box and owns paragraph, leading, color, horizontal left alignment, and the
selected vertical placement. The add-in places that returned box at the padding origin, paints the background and
inside border into the final SVG, and clips it to the authored frame; it performs no second vertical-alignment pass.
It does not use PowerPoint's Fill or Line properties, and it does not ask TeX to draw a full-size outer frame. The
original author source remains unchanged in Alternative Text and the block's declarative style is stored separately
on the shape.
An auto-height block naturally fits its content, so vertical placement normally has no spare height to distribute.
For a fixed-height block, TeX places content at the selected Top, Middle, or Bottom position inside its fixed `vbox`;
the SVG shell only clips any overflow. Line spacing affects ordinary paragraph leading; TeX environments such as
`align` and `gather` retain their own math-row spacing. As in Word, accepting the editor's default style records a real
1.20× TeX leading and an SVG viewport; pre-style default blocks keep their compatible rendering until edited.
Legacy blocks are unchanged merely by opening a document; their first edit, resize, or reflow uses the current
left-aligned TeX-box contract with stable ordinary-text line metrics while preserving their stored vertical placement.

PowerPoint and Word save their profile choices independently.

Fixed Blocks are horizontally left-aligned inside their SVG frame in both hosts. Padding defines the left inset, while
the border is painted inside the outer edge without shrinking the TeX content box; widening a native frame adds space
on the right, while narrowing it clips the right edge.

### Search and accessibility

LaTeX source is stored solely in Alternative Text, which is the semantic description of the visual object. Word's
native Ctrl+F does not search Alternative Text. [Comprehensive Find](https://github.com/zhaiyusci/ComprehensiveFind)
adds a unified search result space for document text and object Alternative Text; LaTeX Blocks deliberately does not
create hidden proxy text.

## Install

The installer is self-contained: it publishes both VSTO add-ins and bundles a compatible private StemTeX runtime.
It requires 64-bit Windows and 64-bit desktop Microsoft Office. Close Word and PowerPoint before installing or
upgrading. See [Releasing](docs/RELEASING.md) to make a package from source.

## Documentation

| Document | Purpose |
| --- | --- |
| [Architecture](docs/ARCHITECTURE.md) | Component boundaries, renderer lifecycle, persistence, and source layout. |
| [Word object model](docs/OBJECT_MODEL.md) | Normative Word representation, geometry, baseline, U+2060, and numbering contract. |
| [PowerPoint scope](docs/POWERPOINT_SCOPE.md) | The intentionally narrower slide-object model. |
| [StemTeX integration](docs/STEMTEX_INTEGRATION.md) | External runtime contract, profile discovery, development stages, and packaging. |
| [Getting started](docs/GETTING_STARTED.md) | Set up, build, register a Debug add-in, and return to the installed product. |
| [Testing](docs/TESTING.md) | Real Word/PowerPoint smoke tests and the PowerPoint width integration test. |
| [Releasing](docs/RELEASING.md) | Versioning, publishing, signing, checksums, and installer verification. |
| [Design decisions](docs/DECISIONS.md) | Deliberate product boundaries and rejected approaches. |
| [Changelog](CHANGELOG.md) | Product-facing changes in each installer version. |
| [Experiments](samples/README.md) | Reproducible host-behavior evidence; not product dependencies. |

## Quick development path

```powershell
pwsh.exe -NoProfile -File .\scripts\Initialize-LaTeXBlocks.ps1
pwsh.exe -NoProfile -File .\scripts\Build-LaTeXBlocks.ps1 -Configuration Debug
pwsh.exe -NoProfile -File .\scripts\Register-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Both
pwsh.exe -NoProfile -File .\scripts\Test-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Both
```

Normal builds and tests never replace the live VSTO registration. Registration is explicit so a Debug build cannot
silently replace the installed add-in. Full prerequisites and host-switching instructions are in
[Getting started](docs/GETTING_STARTED.md).

## Repository layout

```text
src/       Word and PowerPoint VSTO add-ins
tests/     Desktop Office smoke tests and integration checks
scripts/   Build, registration, test, and publishing commands
installer/ Per-user VSTO installer definition
docs/      Product, architecture, workflow, and decision documentation
samples/   Reproducible experiments; not product dependencies
```
