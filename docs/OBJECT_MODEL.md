# LaTeX Block Object Model

## Purpose

An ordinary LaTeX block behaves like one semantic document object, not like an image plus hidden text plus a control
wrapper. The source, visual output, and editing identity occupy three native properties of one Word `InlineShape`.
A numbered display equation adds a same-paragraph Word line-and-number scaffold around that same content object; it
does not add a container or a second copy of the TeX source.

## Representation

| Concern | Word representation | Rule |
| --- | --- | --- |
| Display | Embedded SVG image bytes | Portable output; Word does not need StemTeX to display an existing block. |
| Authoritative source | `InlineShape.AlternativeText` | TeX source only, with Word-stable LF line endings. No AsciiMath, normalized terms, prefixes, or duplicate metadata. |
| Identification | `InlineShape.Title` | Versioned, compact metadata only. |
| Placement | Word `InlineShape` range | Word remains responsible for layout, anchoring, and document persistence. A numbered equation uses same-paragraph manual breaks and tab stops. |
| Equation number | Word `SEQ LaTeXEquation` field | Native searchable text, updated explicitly in document order. |
| Equation reference target | Word bookmark over the `SEQ` result | Name is derived from the stable SVG ID. |

Version 1 metadata is:

```text
LaTeXBlocks/1;id=<guid>;width=<points>;depth=<points>;mode=<auto|fixed>;size=<points>;role=<content|numbered-equation>
```

Word rewrites CRLF in Alternative Text to LF when a DOCX is saved. LaTeX Blocks therefore canonicalizes CRLF and bare
CR to LF before rendering and storage; all other source characters are preserved. This keeps the authoritative source
stable across save/reopen without changing TeX comment boundaries.

The stable ID survives edits. Width is the StemTeX typesetting constraint, not a DPI value and not a raster-image
scale. For a one-line inline source, depth is measured from an invisible dvisvgm marker at the TeX baseline to the
SVG viewport bottom. Metadata written before `role` existed parses as ordinary `content`.

Profile names are discovered from the active StemTeX installation. A directory is offered only when it contains `preamble.tex`. Profile is global add-in state, not object state: changing it affects every subsequent preview, insertion, and rerender. The choice is stored under the current user's LaTeX Blocks settings and is the profile warmed when Word next starts.

Word aligns an inline image through its layout bottom, while `InlineShape.Range.Font.Position` persists only whole points. For a TeX depth `d`, LaTeX Blocks applies `Font.Position = -round(d)` and moves the SVG viewBox by the residual `d - round(d)`. `InlineShapes.AddPicture` also creates a host-only `wp:effectExtent` below an SVG; controlled Word rendering shows that this shifts the inline baseline even though it is not part of the SVG or TeX box. LaTeX Blocks therefore reinserts the same Flat OPC object with `wp:effectExtent.b=0`. `InsertXML` reconstructs the containing paragraph, so the add-in duplicates and restores its complete direct `ParagraphFormat`; editing an SVG must not erase indentation, spacing, or equation tab stops. No numerical effect-extent compensation is mixed into the TeX depth. Fixed-width multiline blocks have no single surrounding-text baseline and retain position zero. A numbered equation is instead one natural-width TeX display box and therefore has one measurable baseline.

Auto-width formulas store the TeX design size used to render them. LaTeX Blocks listens to Word's native Font Size
combo-box command and rerenders formulas in the selected range at the new TeX size. Because Word has no general
formatting-changed event, shortcut and other native formatting paths use a selection-bound snapshot: on entering a
selection, the add-in records each auto formula's image-character size by stable block ID; on leaving, it refreshes
only when that same host size actually changed and the SVG has not already been rendered at the new size. A plain
select/deselect cycle therefore cannot rerender a formula merely because its stored TeX size differed from inherited
Word formatting. No live Word `Range` is retained, and there is no timer or document-wide background scan.

The reference is always the TeX/Western baseline, including when the TeX source contains Chinese. CJK glyphs may extend farther below that line, according to the Chinese font selected by the active StemTeX profile, but they do not redefine it. The add-in never inspects adjacent Word characters, never switches its reference to a Word East Asian font, and never applies a CJK-specific visual offset. Consequently, mixed Chinese/Western content inside the SVG remains internally governed by TeX font metrics, while the SVG's TeX baseline is mapped to Word's running-text baseline exactly once.

Font size is a renderer input, not an image transform. For an auto-width inline object, LaTeX Blocks reads the Word insertion-point size and passes it through StemTeX 0.11's native per-request `font_size_pt` API. StemTeX applies the size inside its live TeX worker, so the resulting SVG, natural width, math metrics, script sizes, optical-size choices, and TeX depth are all produced at the requested size. The add-in does not inject `\fontsize` into the user's source and never enlarges a 10 pt SVG to imitate another TeX size. Fixed-width blocks preserve their explicit document design and editor size.

The Typography control is a one-shot, LaTeX-aware font-size command. It applies the requested size to the selected Word text and rerenders every auto-width LaTeX object in that selection at the same TeX size. This is an explicit user transaction, not a document watcher or timer.

## Layout modes

### Auto-width inline formula

Auto mode is the traditional equation-in-running-text path. The source must be a single-line inline TeX fragment. It is typeset into a TeX `\hbox`; temporary dvisvgm markers report the box's start coordinate, end coordinate, and baseline. LaTeX Blocks then crops the SVG viewport to the natural box width plus the profile's existing 1pt preview border. It removes all measurement markers before embedding the SVG.

The large temporary StemTeX canvas is solely a measurement surface. Its width is not stored as the visible image width and does not create whitespace in Word.

### Fixed-width LaTeX block

Fixed mode preserves the caller's typesetting width and supports display math, multiple lines, and paragraph-like LaTeX content. It deliberately keeps that page-width canvas. Such content does not have one meaningful surrounding-text baseline, so multiline/display blocks remain at Word position zero.

Objects created before the mode field existed parse as `fixed`, preserving their former behavior. Editing can explicitly convert between modes.

### Word-native numbered equation

A numbered equation has one natural-width formula SVG and one Word number. It accepts a collapsed insertion point
inside a paragraph in the main document story; it does not replace a selection or number content in tables, headers,
footnotes, or text boxes. The display line remains part of that logical paragraph:

```text
running text before <manual line break>
<center tab> SVG <right tab> ( SEQ field ) <manual line break>
running text after
```

Either manual break is omitted when the equation is already at that edge of the paragraph. Both breaks are Word line
breaks (`w:br` / character 11), never paragraph marks. The paragraph receives one center tab stop at the midpoint of
the current text column and one right tab stop at the column's right edge. The display line deliberately ignores the
running-text paragraph's left and right indents: Word stores custom tab positions as static offsets from the column's
left edge, so baking an indent into them would leave stale positions after later paragraph-format changes. **Update
Numbers** recomputes these two positions and migrates documents written by the earlier indent-relative implementation.
The center tab aligns the formula independently of the number; the right tab aligns the number. There is no Word table, balancing cell, content
control, separator paragraph, or hidden text. A paragraph containing ordinary tabbed content is rejected because the
equation scaffold must own these two tab stops; another numbered equation in the same paragraph reuses them.
Paragraphs with **Exact** line spacing are rejected because Word cannot expand only the manual-break display line;
Single, At least, and Multiple spacing allow the inline SVG to participate in the line box without rewriting the
paragraph's typography.

The authoritative source may use `\[...\]`, `$$...$$`, `\(...\)`, `$...$`, `displaymath`, or `equation` delimiters.
For rendering only, LaTeX Blocks removes that outer delimiter and submits `\(\displaystyle <body>\)` through the
existing auto-width hbox measurement path. Thus TeX selects display-style fractions, limits, and glyph metrics while
the SVG is cropped to the formula's natural box. The wrapper never enters Alternative Text. A full TeX display
environment would retain the requested page width and is deliberately not embedded as the inline Word object.

The right-aligned tab segment contains literal parentheses around a native field:

```text
( { SEQ LaTeXEquation \\* ARABIC } )
```

The field result alone is bookmarked as `LTXEQ_<32-hex-digit block ID>`. Editing replaces only the SVG and preserves
the block ID, line scaffold, field, and bookmark. The renderer's physical natural width is checked before insertion
or replacement; formulas that would collide with the right-aligned number are rejected before Word is mutated.

The **Update Numbers** command updates only `SEQ LaTeXEquation` fields in the main document story. It is an explicit
operation, not a timer or document-wide background watcher. Moving, copying, or removing complete equation-line
scaffolds can leave cached field results stale until that command or Word's own field-update command runs. Deleting
only the selected SVG is not a semantic equation deletion: it leaves the two tabs, `SEQ` field, and bookmark behind.
The current prototype does not yet expose a dedicated **Delete Equation** command.

One numbered block owns one Word number. `align` and `gather` input can be reduced to one natural-width `aligned` or
`gathered` display box and therefore shares one number. Per-row numbering inside one SVG is deliberately outside this
contract because Word cannot address the internal SVG rows as separate native fields.

## Insert transaction

1. Render the source to SVG with StemTeX.
2. If rendering fails, make no document change.
3. Embed the SVG at the current selection as an `InlineShape`.
4. Write canonical source and metadata to the object.
5. Select the inserted object.

## Edit transaction

1. Read source from Alternative Text and metadata from Title.
2. Render the proposed source before modifying Word.
3. Insert and fully annotate the replacement SVG immediately before the old shape.
4. Delete the old shape only after the replacement is valid.
5. Preserve the block ID while updating width, profile, and source.

A render failure therefore leaves both the previous visual output and its authoritative source untouched.

## Search and accessibility

Alternative Text is semantically appropriate for a visual object whose mathematical or LaTeX meaning is otherwise unavailable to text-oriented tools. Word's native Ctrl+F does not search it. LaTeX Blocks does not work around that limitation with hidden document text because doing so creates a second layout object and a second source of truth.

Comprehensive Find supplies the missing host capability: it searches visible document text and Alternative Text in one result list and selects the visual object for an Alternative Text match. The two add-ins share a document-level contract but do not depend on each other's binaries.

## Deliberate exclusions

- No OMML conversion.
- No hidden proxy runs.
- No content-control wrapper.
- No table or extra paragraph for a numbered equation.
- No automatic per-row Word numbering inside one multiline SVG.
- No background field-renumbering watcher.
- No dedicated numbered-equation deletion command yet; the complete visual-line scaffold is the deletion unit.
- No duplicated plain-language or AsciiMath search representation.
- No multi-object or hidden-run baseline scaffold; inline compensation is one Word whole-point character position plus the fractional residual encoded in the SVG viewBox.
- No mutation of the old object before a successful render.
