# LaTeX Block Object Model

## Purpose

A LaTeX block should behave like one semantic document object, not like an image plus hidden text plus a control wrapper. The source, visual output, and editing identity occupy three native properties of one Word `InlineShape`.

## Representation

| Concern | Word representation | Rule |
| --- | --- | --- |
| Display | Embedded SVG image bytes | Portable output; Word does not need StemTeX to display an existing block. |
| Authoritative source | `InlineShape.AlternativeText` | Exact TeX source only. No AsciiMath, normalized terms, prefixes, or duplicate metadata. |
| Identification | `InlineShape.Title` | Versioned, compact metadata only. |
| Placement | Word `InlineShape` range | Word remains responsible for layout, anchoring, and document persistence. |

Version 1 metadata is:

```text
LaTeXBlocks/1;id=<guid>;width=<points>;depth=<points>;mode=<auto|fixed>
```

The stable ID survives edits. Width is the StemTeX typesetting constraint, not a DPI value and not a raster-image scale. For a one-line inline source, depth is measured from an invisible dvisvgm marker at the TeX baseline to the SVG viewport bottom.

Profile names are discovered from the active StemTeX installation. A directory is offered only when it contains `preamble.tex`. Profile is global add-in state, not object state: changing it affects every subsequent preview, insertion, and rerender. The choice is stored under the current user's LaTeX Blocks settings and is the profile warmed when Word next starts.

Word aligns an inline image through its layout bottom, but an SVG receives an automatic `wp:effectExtent.b=9525` (0.75 pt) below the image extent. `InlineShape.Range.Font.Position` also persists only whole points. For a scaled TeX depth `d`, the effective host depth is therefore `h = d + 0.75 pt`. LaTeX Blocks applies `Font.Position = -round(h)` and moves the SVG viewBox by the residual `h - round(h)`. The two components account for the floating-point TeX depth and Word's host-only effect boundary without adding another Word object or visible marker. Display or multiline blocks have no single surrounding-text baseline and retain position zero.

The reference is always the TeX/Western baseline, including when the TeX source contains Chinese. CJK glyphs may extend farther below that line, according to the Chinese font selected by the active StemTeX profile, but they do not redefine it. The add-in never inspects adjacent Word characters, never switches its reference to a Word East Asian font, and never applies a CJK-specific visual offset. Consequently, mixed Chinese/Western content inside the SVG remains internally governed by TeX font metrics, while the SVG's TeX baseline is mapped to Word's running-text baseline exactly once.

The maintained StemTeX profiles use the 10 pt `article` base size. An auto-width inline object is therefore scaled uniformly by `Word insertion-point font size / 10 pt`; its measured TeX depth is scaled by the same ratio before Word positioning. This makes 10 pt TeX text become 11 pt beside ordinary 11 pt Word text without changing the baseline definition. Fixed-width blocks preserve their explicit canvas size and are not scaled from surrounding text.

## Layout modes

### Auto-width inline formula

Auto mode is the traditional equation-in-running-text path. The source must be a single-line inline TeX fragment. It is typeset into a TeX `\hbox`; temporary dvisvgm markers report the box's start coordinate, end coordinate, and baseline. LaTeX Blocks then crops the SVG viewport to the natural box width plus the profile's existing 1pt preview border. It removes all measurement markers before embedding the SVG.

The large temporary StemTeX canvas is solely a measurement surface. Its width is not stored as the visible image width and does not create whitespace in Word.

### Fixed-width LaTeX block

Fixed mode preserves the caller's typesetting width and supports display math, multiple lines, and paragraph-like LaTeX content. It deliberately keeps that page-width canvas. Such content does not have one meaningful surrounding-text baseline, so multiline/display blocks remain at Word position zero.

Objects created before the mode field existed parse as `fixed`, preserving their former behavior. Editing can explicitly convert between modes.

## Insert transaction

1. Render the source to SVG with StemTeX.
2. If rendering fails, make no document change.
3. Embed the SVG at the current selection as an `InlineShape`.
4. Write exact source and metadata to the object.
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
- No duplicated plain-language or AsciiMath search representation.
- No multi-object or hidden-run baseline scaffold; inline compensation is one Word whole-point character position plus the fractional residual encoded in the SVG viewBox.
- No mutation of the old object before a successful render.
