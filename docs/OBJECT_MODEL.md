# LaTeX Block Object Model

This document specifies the Word host. PowerPoint intentionally uses only the free-standing block subset described in
[POWERPOINT_SCOPE.md](POWERPOINT_SCOPE.md); none of the Word inline-formula rules below apply to PowerPoint.

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

For a fixed-width block, the saved `width` remains an absolute TeX layout width in points. The primary Office controls
therefore use exact points as well: 30–450 pt, a 0.5 pt step, one decimal place, and a 360 pt default, matching the
StemTeX GUI. It is not a live binding to future page or slide size changes. PowerPoint stores visual
scale separately from the SVG's intrinsic dimensions. A horizontal-only shape resize changes layout width and
rerenders after PowerPoint's resize-completed event, while a resize with a vertical component changes visual scale
and reuses the SVG. This is event-driven direct manipulation, not geometry polling.

Word resolves its actual container (usable table-cell width, text-frame width after internal margins, or the active
section column's own width) for placement and available geometry, but a new fixed block begins at the independent
360 pt typesetting width. Its metadata stores that absolute point value, so moving the object later does not silently
reflow it.

Profile names are discovered from the active StemTeX installation. A directory is offered only when it contains `preamble.tex`. Profile is global add-in state, not object state: changing it affects every subsequent preview, insertion, and rerender. The choice is stored under the current user's LaTeX Blocks settings and is the profile warmed when Word next starts.

Word aligns an inline image through its layout bottom, while `InlineShape.Range.Font.Position` persists only whole points. For a surrounding run position `p` and TeX depth `d`, LaTeX Blocks applies `Font.Position = p - round(d)` and moves the SVG viewBox by the residual `d - round(d)`. `InlineShapes.AddPicture` initially quantizes an SVG's physical dimensions through CSS pixels. LaTeX Blocks therefore converts the final SVG dimensions directly to EMUs (`cx = round(width_pt × 12700)`, `cy = round(height_pt × 12700)`) and writes the exact pair to both `wp:inline/wp:extent` and `pic:spPr/a:xfrm/a:ext`. It also normalizes the host-only `wp:effectExtent.b` to zero. These are vector-layout coordinates, not a DPI calculation. `InsertXML` reconstructs the containing paragraph, so the add-in duplicates and restores its complete direct `ParagraphFormat`; editing an SVG must not erase indentation, spacing, or equation tab stops. No numerical effect-extent compensation is mixed into the TeX depth. Fixed-width multiline blocks have no single surrounding-text baseline and retain position zero. A numbered equation is instead one natural-width TeX display box and therefore has one measurable baseline.

The negative position belongs only to the inline picture character. Word omits that character-level `w:position` and `w:sz` when the one-character image range is exported as Flat OPC, so the add-in must restore both the TeX design size and baseline position on the final image after effect-extent normalization. `w:sz` is semantic host-run state: changing it does not rescale the SVG, but it keeps Word's Font Size UI and formula refresh logic consistent with the size StemTeX actually rendered. The add-in then collapses the selection after the object instead of leaving the compensated picture character selected. At that boundary Word natively resolves the caret from the following running-text run (or the paragraph insertion format), so normal continuation text cannot inherit the picture's `w:position` or `w:noProof`.

Word expands a U+0020 immediately adjacent to an `InlineShape` toward a host-chosen half-em advance. The excess is font-dependent: it is close to one additional normal space in Times New Roman, larger in Aptos and Calibri, and nearly zero for fonts whose ordinary space is already wide. It is independent of the SVG width, DPI, drawing-run font, East Asian auto-spacing, and character grid. LaTeX Blocks does not replace, scale, or negatively space the user's U+0020. Before insertion it measures the ordinary advance for each adjacent space. A lone U+0020 can be measured directly. Two U+0020 characters that already touch are not a valid direct reference because Word expands their reported geometry even before an image exists; in that case the add-in uses an equivalently formatted isolated space on the same visual line, or measures the same Word font formatting in a temporary hidden document and closes it immediately. After insertion it measures the formatted character beside the SVG and writes only the positive excess to the corresponding signed `wp:effectExtent.l` or `.r`: `e = -round((inlineAdvance - naturalAdvance) × 12700)`. A negative left extent moves the SVG canvas leading edge and the following text left by that amount; a negative right extent leaves the canvas origin fixed and shortens only its trailing advance. Therefore the page-coordinate postcondition is authoritative: the canvas leading edge must follow the preceding word by exactly the natural space, and following punctuation or a following natural space must begin at the corresponding SVG edge. After the first normalization, the add-in re-derives the extents from the completed DrawingML object and performs one synchronous corrective rewrite only if this final answer differs; it never waits, polls, or watches the document. On edit, an equivalent current same-line space is authoritative; the persisted natural width and scratch measurement are fallbacks and a guard against one Word layout quantum of measurement noise. A real font or size change therefore produces a fresh extent. The SVG's exact physical dimensions and viewBox remain unchanged. A side with no ordinary adjacent space keeps a zero effect extent. The values are recomputed as part of an explicit insert, edit, or size refresh and persist in DOCX; there is no spacing watcher for text typed later.

Auto-width formulas store the TeX design size used to render them. LaTeX Blocks listens to Word's native Font Size
combo-box command and rerenders formulas in the selected range at the new TeX size. Because Word has no general
formatting-changed event, shortcut and other native formatting paths use a selection-bound snapshot: on entering a
selection, the add-in records each auto formula's image-character size by stable block ID; on leaving, it refreshes
only when that same host size actually changed and the SVG has not already been rendered at the new size. A plain
select/deselect cycle therefore cannot rerender a formula merely because its stored TeX size differed from inherited
Word formatting. No live Word `Range` is retained, and there is no timer or document-wide background scan.

The reference is always the TeX/Western baseline, including when the TeX source contains Chinese. CJK glyphs may extend farther below that line, according to the Chinese font selected by the active StemTeX profile, but they do not redefine it. Baseline resolution never inspects adjacent Word characters, never switches its reference to a Word East Asian font, and never applies a CJK-specific visual offset. Consequently, mixed Chinese/Western content inside the SVG remains internally governed by TeX font metrics, while the SVG's TeX baseline is mapped to Word's running-text baseline exactly once.

Font size is a renderer input, not an image transform. For an auto-width inline object, LaTeX Blocks reads `Selection.Font.Size`, which is Word's actual typing size. This distinction matters at a run boundary and after changing the size of a collapsed caret: `Selection.Range.Font.Size` can still describe the character to the right even though newly typed text uses another size. For a mixed non-collapsed selection, Word reports `wdUndefined`; replacement then follows Word's native rule and uses the first selected character's insertion size. The resolved size is passed through StemTeX 0.12's native per-request `font_size_pt` API. StemTeX applies the size inside its live TeX worker, so the resulting SVG, natural width, math metrics, script sizes, optical-size choices, and TeX depth are all produced at the requested size. The add-in does not inject `\fontsize` into the user's source and never enlarges a 10 pt SVG to imitate another TeX size. Fixed-width blocks preserve their explicit document design and editor size.

The Typography control is a one-shot, LaTeX-aware font-size command. It applies the requested size to the selected Word text and rerenders every auto-width LaTeX object in that selection at the same TeX size. This is an explicit user transaction, not a document watcher or timer.

## Layout modes

### Auto-width inline formula

Auto mode is the traditional equation-in-running-text path. The source must be a single-line inline TeX fragment. It is typeset into a TeX `\hbox`; temporary dvisvgm markers report the box's start coordinate, end coordinate, and baseline. Wrapper line breaks are suppressed so they cannot become TeX interword glue at either edge; explicit spacing written in the source is preserved. Horizontally, LaTeX Blocks removes the profile's generic 1pt page-preview border and crops to the union of the logical TeX box and dvisvgm's actual ink bounds, plus 0.05pt vector safety. Real glyph, accent, rule, and operator overhangs therefore remain visible without adding fixed padding to ordinary formulas. It removes all measurement markers before embedding the SVG.

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
5. Normalize exact drawing dimensions and host-only effect extents, then place the caret after the object.

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
