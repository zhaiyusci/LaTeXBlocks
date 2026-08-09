# Word Object Model

This is the normative representation and layout contract for the Word host. PowerPoint uses only a separate
free-standing block model; none of the Word inline-formula rules below apply there. See
[POWERPOINT_SCOPE.md](POWERPOINT_SCOPE.md) and [ARCHITECTURE.md](ARCHITECTURE.md) for the host boundary.

## Purpose

An ordinary LaTeX block behaves like one semantic document object, not like an image plus hidden text plus a control
wrapper. The source, visual output, and editing identity occupy native properties of one Word drawing object. It is
normally an `InlineShape`; a user may make a fixed-width content block floating, in which case Word losslessly changes
it into a `Shape`. A numbered display equation adds a same-paragraph Word line-and-number scaffold around its inline
content object; it does not add a container or a second copy of the TeX source.

## Representation

| Concern | Word representation | Rule |
| --- | --- | --- |
| Display | Embedded SVG image bytes | Portable output; Word does not need StemTeX to display an existing block. |
| Authoritative source | `InlineShape.AlternativeText` or floating `Shape.AlternativeText` | TeX source only, with Word-stable LF line endings. No AsciiMath, normalized terms, prefixes, or duplicate metadata. |
| Identification | `InlineShape.Title` or floating `Shape.Title` | Versioned, compact metadata only. |
| Placement | Word `InlineShape` range; optionally a floating `Shape` for fixed content | Word remains responsible for layout, anchoring, and document persistence. An auto-width content formula alone has one U+2060 WORD JOINER immediately before and after its image; a numbered equation uses same-paragraph manual breaks and tab stops. |
| Equation number | Word `SEQ LaTeXBlockEq` field | Native searchable text, updated explicitly in document order. The matching Word Caption Label declares the category, while the stable bookmark remains the individual target. |
| Equation reference target | Word bookmark over the `SEQ` result | Name is derived from the stable SVG ID. |

Version 1 metadata is:

```text
LaTeXBlocks/1;id=<guid>;width=<points>;depth=<points>;mode=<auto|fixed>;size=<points>;role=<content|numbered-equation>[;framewidth=<points>;frameheight=<points>;style=<compact-style>]
```

Word rewrites CRLF in Alternative Text to LF when a DOCX is saved. LaTeX Blocks therefore canonicalizes CRLF and bare
CR to LF before rendering and storage; all other source characters are preserved. This keeps the authoritative source
stable across save/reopen without changing TeX comment boundaries.

The stable ID survives edits. Width is the StemTeX typesetting constraint, not a DPI value and not a raster-image
scale. For a natural-width single-baseline source, depth is measured from an invisible dvisvgm marker at the TeX
baseline to the SVG viewport bottom. It includes the standard minimum `\strutbox` depth, or a larger natural content
depth when required. `framewidth` and `frameheight`, when present, are the authored physical SVG root for a floating
fixed block; they are deliberately distinct from the TeX measure. `style`, when present, is a delimiter-safe compact
serialization of the shared Block style (leading, padding, vertical placement, text/fill/border colors, and border
width). It appears only for a fixed ordinary Content Block; it never replaces TeX source in Alternative Text.
The version-one Word style payload's fourth `t/m/b` slot stores Top/Middle/Bottom vertical placement. Current writers
emit the selected value, and readers restore all three values; retaining this version-one layout keeps existing Blocks
compatible without a metadata-version change.
Metadata written before `role`, frame fields, or `style` continues to parse as ordinary `content` with the historical
unstyled rendering route.

For a fixed-width block, the saved `width` remains an absolute TeX layout width in points. The primary Word controls
therefore use exact points as well: 30–2000 pt, a 0.5 pt step, one decimal place, and a 360 pt default. The wider
range lets a reflowed floating Word frame retain its user-selected geometry. It is not a live binding to future page
or section size changes.

Word resolves its actual container (usable table-cell width, text-frame width after internal margins, or the active
section column's own width) for placement and available geometry, but a new fixed block begins at the independent
360 pt typesetting width. Its metadata stores that absolute point value, so moving the object later does not silently
reflow it.

A fixed-width content block has the same frame contract whether Word represents it as an `InlineShape` (**In Line with
Text**) or as a floating `Shape` under any Layout Option. `width` is the TeX typesetting measure, while
`framewidth`/`frameheight` are the user-owned outer SVG frame. A size change creates a fresh TeX render; the root SVG
is then reframed to the exact outer dimensions by expanding its viewport (transparent space) or reducing it (clip),
never by scaling TeX glyph coordinates. For a styled Block, the fresh root also paints padding, fill, and four
in-viewport border strips; the border therefore always remains at the current frame edge after a resize. The outer
frame minus padding is the exact TeX content box: TeX keeps its contents horizontally left-aligned and performs the
stored Top/Middle/Bottom placement inside a fixed-height `vbox`. The SVG shell places that box at the padding origin
and does not calculate another vertical offset. A changed frame width derives a new TeX measure from the prior SVG
root; a height-only change rerenders at the existing measure.

Word's COM model has no `AfterShapeSizeChange` equivalent. The add-in therefore observes Word's documented,
process-scoped native mouse-capture completion event, posts one UI-thread turn after mouse-up, and then reads the final
Block geometry. The event callback never accesses Word COM and it is not a geometry poll, document watcher, window
subclass, or global mouse hook. `WindowSelectionChange` remains a fallback for a non-mouse resize or an environment
where the operating system refuses the monitor. **Reflow Frame** performs the same asynchronous operation explicitly
for the currently selected fixed Block. The comparison is between the frame at the start and end of one gesture,
not against stale persisted metadata, so translation and rotation do not request a render and rapid consecutive
resizes compose correctly. The native drag may be
temporarily shown as a Word picture transform, but the persisted replacement has no such scale transform. For a
floating replacement the add-in temporarily returns the shape to the inline representation, then restores its relative
reference frame, left/top position, wrapping type/side/distances, overlap setting, layout-in-cell, anchor lock, and
rotation. Editing source or TeX measure keeps the same outer frame in either representation. This option applies only
to fixed-width ordinary content. Auto-width formulas remain inline for their U+2060 word-spacing scaffold, and
numbered equations remain inline for their tab/`SEQ`/bookmark layout.

The Word frame is independent of the fixed Block's 30–2000 pt TeX measure policy. A valid positive native frame is
stored and emitted at its exact physical SVG extent, including an extent above 2000 pt; at those extremes the TeX
measure remains bounded, so the resulting frame may intentionally contain extra space or clip content.

Profile names are discovered from the active StemTeX installation. A directory is offered only when it contains
`preamble.tex`. Profile is Word-host state, not object state: changing it affects every subsequent Word preview,
insertion, and rerender. Word persists its preference independently from PowerPoint and warms that profile when it
next starts. Runtime selection and profile discovery are specified in [STEMTEX_INTEGRATION.md](STEMTEX_INTEGRATION.md).

## Geometry and vertical baseline

Word aligns an inline image through its layout bottom, while `InlineShape.Range.Font.Position` persists only whole points. For TeX depth `d`, LaTeX Blocks applies `Font.Position = -round(d)` and moves the SVG viewBox by the residual `d - round(d)`. `Font.Position` is already relative to the current Word line baseline; neighboring text positions therefore never enter formula placement. `InlineShapes.AddPicture` initially quantizes an SVG's physical dimensions through CSS pixels. LaTeX Blocks therefore converts the final SVG dimensions directly to EMUs (`cx = round(width_pt × 12700)`, `cy = round(height_pt × 12700)`) and writes the exact pair to both `wp:inline/wp:extent` and `pic:spPr/a:xfrm/a:ext`. Every `wp:effectExtent` side (`l`, `t`, `r`, and `b`) is normalized to zero. These are vector-layout coordinates, not a DPI calculation. For auto-width content, the SVG viewport is the exact TeX box: explicit TeX glue remains part of that box, while the add-in supplies no display padding, ink padding, or horizontal-space correction. `InsertXML` reconstructs the containing paragraph, so the add-in duplicates and restores its complete direct `ParagraphFormat`; editing an SVG must not erase indentation, spacing, or equation tab stops. Fixed-width multiline blocks have no single text baseline and retain position zero. A numbered equation is instead one natural-width TeX display box and therefore has one measurable baseline.

The negative position belongs only to the inline picture character. Word omits that character-level `w:position` and
`w:sz` when the one-character image range is exported as Flat OPC, so the add-in restores both the TeX design size and
baseline position on the final image after normalizing its drawing extent. `w:sz` is semantic host-run state:
changing it does not rescale the SVG, but it keeps Word's Font Size UI and formula refresh logic consistent with the
size StemTeX actually rendered.

## Word character-format ownership

Replacing an SVG also replaces its one-character Word run. Updates therefore merge properties by ownership instead
of accepting Word's defaults or copying the old run indiscriminately:

| Property group | Update rule |
| --- | --- |
| Authoritative formula state | TeX source, stable ID, role, layout mode, typesetting width, explicit Block style, and authored frame data come from the edit request plus existing object metadata. They are not inferred from character formatting. |
| Native renderer inputs | For Auto and numbered formulas, `Font.Size`/`SizeBi` represent one TeX design size; native Font Size changes rerender the SVG. A styled Fixed Block takes its TeX size from durable editor metadata instead. |
| External foreground paint | Word Font Color, Word Graphics Fill, and PowerPoint Graphics Fill are host paint operations for every formula kind, including numbered equations and Fixed Blocks. They update the Office Graphics Fill without invoking StemTeX or replacing SVG media. Word retains its native colour descriptor—Automatic, direct BGR, or theme slot plus tint/shade—alongside the resolved display fill. Explicit colours authored inside TeX remain child-level SVG overrides. |
| Formula-derived output | SVG bytes and physical extents, TeX depth, and frame geometry come from the new render. Auto/numbered `Font.Position` is recomputed as `-round(new depth)` relative to Word's current line baseline; Fixed Content is reset to position zero. `Subscript` and `Superscript` are cleared because script placement belongs in TeX source, not in a second Word transform. |
| Independent Word run format | Font family metadata, bold/italic/underline/strike/hidden state, spacing, scaling, kerning, proofing state, highlight, and language IDs survive replacement. Except for the renderer inputs listed above, these values do not alter glyphs inside the SVG; spacing/scaling may still alter Word's treatment of the drawing character around its exact SVG extent. The resolved direct values are preserved, not the identity of a Word character style. A native theme colour is the exception: its theme slot and tint/shade are deliberately restored after drawing normalization rather than downgraded to RGB. |
| Paragraph and UI state | Direct `ParagraphFormat` and document-owned paragraph marks survive normalization. An exact InlineShape selection transfers to the replacement only if that same old object is still selected when rendering commits. A mixed text range uses one shared duplicated-range lease and is restored after each selected formula replacement only while the user still owns that range; a later caret move is never reversed. |

For a multi-formula Font Size action, the renderer inputs and derived outputs change per formula, while every
independent run-format and Graphics Fill value is captured and replayed per formula. Formulas in different paragraphs
are replaced through one Flat OPC envelope for their common Word story; paragraph marks and intervening text remain
inside that envelope unchanged. Only ordinary Auto Content formulas share this envelope. Numbered equations and Fixed
Blocks retain their distinct placement/frame contracts and are routed through their individual update paths even when
the same Word selection also contains batchable formulas. The asynchronous drawing replacement is one custom Undo record. It deliberately does
not absorb Word's already-completed native Font Size command into that record, so Undo first restores the old formula
drawings as one batch and a subsequent Undo remains Word's own formatting action.

The commit path rechecks the live native size and colour before replacing the object. If either changed while StemTeX
was rendering, the stale SVG is discarded and one merged refresh is queued. Word enables Font Color for an exact
`InlineShape` selection but its built-in command does not modify that drawing run. LaTeX Blocks therefore listens to
one value-free native-format transaction stream (`Began`, `Committed`, `Canceled`) rather than exposing UIA/MSAA
details to the document layer. The main button has a direct Invoke. A gallery popup is a separate `NetUIToolWindow`:
MSAA selection identifies the active hovered swatch, but hover alone never commits; a left-button down/up pair that
starts on that same live popup swatch (or a provider Invoke on builds that expose one) confirms the gesture. The
candidate survives the popup's short close-ordering window, but an Escape followed by a click in its stale screen
rectangle cannot commit. A generation-bound timer defers the semantic commit until Word has processed the native
mouse command. Every lifecycle signal then leaves one FIFO UI-thread queue, so `Began` always precedes the matching
terminal signal even when UIA, MSAA, and Win32 callbacks arrive on different threads. More Colors commits only after
**OK** and dialog close; opening a menu, Cancel, Escape, and window close do not commit.

For an exact formula selection, a confirmed action uses a collapsed-caret `ExecuteMso("FontColorPicker")` transaction
to read Word's current picker descriptor, immediately restores the exact picture selection, applies the descriptor to
the drawing run, and queues the render. For an ordinary range containing text and one or more formulas, Word has
already written the native colour to every selected drawing character; the transaction coordinator reads each frozen
formula target and queues only its colour delta, without probing or rewriting ordinary text. Programmatic
accessibility echoes are suppressed. These paths are separate from the generic mouse-capture monitor used only for
frame resizing; they do not poll and never make the U+2060 scaffold the user's selection.

The update then snapshots independent run format, inserts and normalizes the new drawing, reapplies that format,
overwrites renderer inputs and derived placement, and finally reapplies any native theme descriptor. Consequently
changing one supported property cannot restore the others to defaults, while a damaged old baseline can never be
mistaken for user-owned state.

## Inline word-spacing scaffold

Word can change a U+0020 immediately adjacent to an `InlineShape`: most fonts expand it toward a host-chosen half-em advance, while SimSun can narrow it. This is host layout behavior, independent of SVG width, DPI, drawing-run font, East Asian auto-spacing, and character grid. For an auto-width ordinary-content formula, LaTeX Blocks avoids that direct adjacency by placing one U+2060 WORD JOINER immediately before the image and one immediately after it. The user's U+0020 characters are neither rewritten nor measured, and the SVG carries neither host spacing nor compensating geometry. The boundary pair belongs to the placement scaffold, not to the formula source or title metadata: the authoritative TeX source remains exclusively in `InlineShape.AlternativeText`.

The pair is an exact invariant, not an insertion-time convenience. On a fresh insertion, LaTeX Blocks creates one
immediately adjacent joiner on each side. On an edit, it reuses an immediately adjacent existing joiner instead of
inserting another one, so repeated updates preserve exactly one leading and one trailing U+2060. If an auto formula
is changed into a fixed-width block, its unshared boundary joiners are removed; a joiner shared with an immediately
neighboring auto formula remains for that neighbor. Fixed-width blocks and Word-native numbered equations are not
wrapped because they do not participate in running-text word spacing. After insertion or replacement the selection
collapses after the trailing joiner, so following typing is placed outside the formula scaffold and cannot inherit the
picture character's `w:position` or `w:noProof`. When the user selects the formula itself, Word's exact InlineShape
selection is retained; the boundary pair is never permanently included merely to route a native formatting command.

## Font-size synchronization

Auto-width formulas store the TeX design size used to render them. LaTeX Blocks listens to Word's native Font Size
combo-box command and rerenders formulas in the selected range at the new TeX size. Because Word has no general
formatting-changed event, native formatting paths use a selection-bound snapshot: on entering a selection, the add-in
records each auto formula's image-character size and colour against that selected drawing identity. The native Font
Size combo has its own command callback; Font Color uses the dedicated transaction described above and reconciles a
mixed range immediately. Keyboard, macro, and other formatting paths are still checked when the user leaves the
selection. A refresh occurs only when the host value actually changed and the SVG has not already been rendered at
that value. A plain select/deselect cycle therefore cannot rerender a formula merely because its stored TeX size
differed from inherited Word formatting. The coordinator never retains the mutable global `Selection`; a bounded
duplicated-range lease exists only for the affected asynchronous replacements. There is no recurring timer or
document-wide background scan.

Font size is a renderer input, not an image transform. For an auto-width inline object, LaTeX Blocks reads `Selection.Font.Size`, which is Word's actual typing size. This distinction matters at a run boundary and after changing the size of a collapsed caret: `Selection.Range.Font.Size` can still describe the character to the right even though newly typed text uses another size. For a mixed non-collapsed selection, Word reports `wdUndefined`; replacement then follows Word's native rule and uses the first selected character's insertion size. The resolved size is passed through StemTeX 0.12's native per-request `font_size_pt` API. StemTeX applies the size inside its live TeX worker, so the resulting SVG, natural width, math metrics, script sizes, optical-size choices, and TeX depth are all produced at the requested size. The add-in does not inject `\fontsize` into the user's source and never enlarges a 10 pt SVG to imitate another TeX size. Fixed-width blocks preserve their explicit document design and editor size.

The Typography control is a one-shot, LaTeX-aware font-size command. It applies the requested size to the selected Word text and rerenders every auto-width LaTeX object in that selection at the same TeX size. This is an explicit user transaction, not a document watcher or timer.

## CJK baseline policy

The reference is always the TeX/Western baseline, including when the TeX source contains Chinese. CJK glyphs may
extend farther below that line, according to the Chinese font selected by the active StemTeX profile, but they do not
redefine it. Baseline resolution never inspects adjacent Word characters, never switches its reference to a Word East
Asian font, and never applies a CJK-specific visual offset. Consequently, mixed Chinese/Western content inside the SVG
remains internally governed by TeX font metrics, while the SVG's TeX baseline is mapped to Word's running-text
baseline exactly once.

## Layout modes

### Natural-size formula (Auto)

Auto mode is the natural-size formula path. It covers both inline math and unnumbered display math imported from LaTeX. Display math differs by using `\displaystyle` and Word paragraph placement; it does not acquire a fixed-width SVG canvas. The rendered fragment is typeset into a TeX `\hbox` containing one zero-width standard `\strut`; temporary dvisvgm markers report the box's start coordinate, end coordinate, and baseline. The strut supplies a stable minimum `0.7 × \baselineskip` height and `0.3 × \baselineskip` depth without changing the formula width, while taller content expands naturally. Wrapper line breaks are suppressed so they cannot become TeX interword glue at either edge; explicit spacing written in the source is preserved. The embedded SVG uses `\PreviewBorder=0pt`. Its horizontal viewport is the union of the logical TeX width and genuine ink overhang, while its vertical viewport is the completed TeX line box; no vector safety padding or application-added edge spacing is introduced. All temporary measurement markers are removed before embedding the SVG.

An auto-width formula whose role is ordinary content is surrounded by exactly one U+2060 WORD JOINER on each side.
Those invisible boundary characters keep any user-authored U+0020 spaces out of Word's image-adjacency layout path;
they do not contain source, do not receive object metadata, and are reused rather than duplicated when the formula is
updated. The selection ends after the trailing joiner. This scaffold is deliberately absent from fixed-width blocks
and numbered equations.

The large temporary StemTeX canvas is solely a measurement surface. Its width is not stored as the visible image width and does not create whitespace in Word.

### Fixed-width LaTeX block

Fixed mode is reserved for an explicit user-sized Block. It preserves the caller's outer frame and supports display math, multiple lines, and paragraph-like LaTeX content inside that frame. It deliberately keeps the requested typesetting canvas. Such content does not have one meaningful surrounding-text baseline, so multiline Blocks remain at Word position zero. Its editor exposes TeX font size, line spacing, uniform padding, Top/Middle/Bottom vertical placement, text color, fill, and border controls. The authored outer frame minus padding becomes an exact TeX box with zero paragraph indentation. TeX owns leading, text color, horizontal left alignment, and vertical placement inside its fixed-height `vbox`. Ordinary text receives stable first/final line height and depth so x-height-only content cannot touch a Top-aligned content edge; standalone display math remains unwrapped. The SVG shell only adds padding, fill, an inside border, and clipping; it neither aligns the content again nor uses a Word Shape Fill/Line or TeX `\fbox`.

Objects created before the mode field existed parse as `fixed`, preserving their former behavior. An inline fixed
Content Block can be explicitly converted between modes. A floating Block remains fixed while it owns a physical
Word frame; convert it back to **In Line with Text** before changing it to an auto-width formula.

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
the resulting single-baseline box receives the same standard minimum strut as an inline formula. The wrapper never
enters Alternative Text. A full TeX display
environment would retain the requested page width and is deliberately not embedded as the inline Word object.

The right-aligned tab segment contains literal parentheses around a native field:

```text
( { SEQ LaTeXBlockEq \\* ARABIC } )
```

The field result alone is bookmarked as `LTXEQ_<32-hex-digit block ID>`. Editing replaces only the SVG and preserves
the block ID, line scaffold, field, and bookmark. The renderer's physical natural width is checked before insertion
or replacement; formulas that would collide with the right-aligned number are rejected before Word is mutated.

`LaTeXBlockEq` is also registered as a Word Caption Label when the add-in starts, so Word recognizes the field
identifier as one explicit equation category. That registration is Word application state, not DOCX data, and is only
a UI convenience. Word's native Cross-reference dialog does not treat this tab-scaffold's bare `SEQ` fields as its
own caption objects, so it cannot serve as the equation picker. This is not a restriction on reference precision:
Word `REF` fields can target the individual `LTXEQ_...` bookmark even when several visual equation lines share one
paragraph. The add-in's **Equation Reference** picker enumerates those verified bookmarks and inserts native
`REF LTXEQ_<id> \\h` fields, surrounded by ordinary parentheses, rather than relying on the built-in dialog.

The **Update Numbers** command updates `SEQ LaTeXBlockEq` fields and then LaTeX Blocks' matching `REF` fields in the
main document story. It also migrates the former private `SEQ LaTeXEquation` identifier into this public category,
so an upgraded document retains one numbering sequence. Word Caption Labels are application state rather than DOCX
data; the stable document-level reference target remains the bookmark around the field result. The command is an
explicit operation, not a timer or document-wide background watcher. The picker and update pass deliberately apply
only to the main document story; headers, footnotes, comments, and text boxes are outside this first
cross-reference scope. Moving, copying, or removing complete equation-line
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
5. Normalize exact drawing dimensions and set all `wp:effectExtent` sides to zero. For auto-width content only,
   ensure the single leading and trailing U+2060 boundary joiners, then place the caret after the trailing joiner.

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
- No multi-object or hidden-run baseline scaffold; vertical baseline mapping is one Word whole-point character position plus the fractional residual encoded in the SVG viewBox. Horizontal add-in spacing remains zero.
- No mutation of the old object before a successful render.
