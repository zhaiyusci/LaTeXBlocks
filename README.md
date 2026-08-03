# LaTeX Blocks

LaTeX Blocks puts editable LaTeX into Microsoft Office as portable SVG objects. StemTeX performs TeX layout and
SVG rendering; the Office add-in owns placement, editing, and persistence. The source of each block remains on the
visual object itself, so a document continues to display correctly without the renderer installed.

The project contains separate Word and PowerPoint VSTO add-ins. StemTeX is an external infrastructure project, not a
Git submodule of this repository; see [StemTeX integration](docs/STEMTEX_INTEGRATION.md).

## Host capabilities

| Capability | Word | PowerPoint |
| --- | --- | --- |
| Traditional inline formula | Yes — exact TeX SVG box, TeX baseline mapping, and a U+2060 boundary scaffold | No |
| Fixed-width LaTeX block | Yes | Yes — positioned slide shape |
| Numbered display equation | Yes — Word-native `SEQ` field and tab stops | No |
| Edit source on the visual object | Yes | Yes |
| Source storage | SVG `AlternativeText` | Shape `AlternativeText` |
| Host font-size integration | Selection-aware inline refresh | Snapshot text size at insertion; explicit block-size control thereafter |

Word auto-width content formulas are the only objects surrounded by U+2060 WORD JOINER characters. This keeps a
user-authored ordinary space out of Word's special inline-image spacing path without putting padding or negative
spacing into the TeX SVG. Fixed blocks and numbered equations are not wrapped. The full invariant is specified in
the [Word object model](docs/OBJECT_MODEL.md).

PowerPoint intentionally supports only free-standing blocks. It does not attempt to make an SVG participate in a
text run or simulate inline baseline behavior; see [PowerPoint scope](docs/POWERPOINT_SCOPE.md).

## Use

### Word

The **LaTeX Blocks** Ribbon tab provides **Insert Inline Formula**, **Insert LaTeX Block**, **Insert Numbered
Equation**, and **Edit LaTeX Block**. Select a recognized block to edit its authoritative source. The selected
profile is a preference of the Word host, not of an individual block.

Numbered equations use Word fields. After moving, copying, or deleting complete numbered-equation lines, run
**Update Numbers** (or Word's own field-update command). Deleting only the SVG is not a semantic equation deletion:
the Word field and tab scaffold remain.

### PowerPoint

Use **Insert Block** or **Edit Block**. A new block starts with the ordinary text size at the active caret/selection
when one exists. **Typesetting width (pt)** is an explicit TeX layout control; **TeX size (pt)** rerenders at a
different design size. Every native PowerPoint *size* change also queues a real TeX layout pass: a changed width
derives a new typesetting measure, while a height-only change rerenders at the current measure under the new frame.
Moving or rotating the shape does not render. SVG artwork is never visually scaled or cropped. If fixed-size TeX
cannot satisfy a constrained dimension after reflow, the frame retains the natural safe extent on that axis. Use
**TeX size (pt)** when the formula itself should become larger or smaller.

PowerPoint and Word save their profile choices independently.

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
