# Design Decisions

This document records the product boundaries established by experiments and host behavior. It is not a roadmap; the
items below describe intentional current choices.

## Render TeX as SVG through StemTeX

LaTeX Blocks uses StemTeX's hot TeX renderer and embeds self-contained SVG. It does not convert LaTeX into OMML or
depend on MathType. SVG preserves the TeX result and keeps rendering separate from Office's incomplete math layout.

## One visual object, one source

The authoritative TeX source is Alternative Text on the SVG object. The project does not add hidden proxy runs,
ASCII-math copies, normalized search vocabulary, content-control wrappers, or a second source record. This keeps the
document model truthful. [Comprehensive Find](https://github.com/zhaiyusci/ComprehensiveFind) supplies the separate
capability Word lacks: searching Alternative Text together with ordinary document text.

## Word inline formulas use an exact TeX box plus U+2060 boundaries

Word changes the advance of a literal U+0020 space immediately next to an `InlineShape`. The add-in does not put
padding, negative `\hspace`, character scaling, or signed `wp:effectExtent` compensation into the formula. Instead,
an auto-width ordinary-content formula receives one U+2060 WORD JOINER immediately on each side. That removes direct
image adjacency while preserving the user's U+0020 characters and the exact TeX SVG box. The pair is reused on edit;
fixed blocks and numbered equations do not use it. See [OBJECT_MODEL.md](OBJECT_MODEL.md).

## Vertical baseline mapping is distinct from word spacing

Word persists character position only in whole points. The add-in maps TeX depth to the picture character's Word
baseline position and stores the fractional residual in the SVG viewBox. This is vertical TeX-box geometry, not a
horizontal word-spacing workaround.

## PowerPoint has blocks, not inline math

PowerPoint's text system is not a full rich-text layout surface for arbitrary embedded OLE/SVG inline objects. The
PowerPoint add-in therefore creates only free-standing LaTeX Blocks. It inherits surrounding text size at insertion
but does not claim a text-run baseline or mutate neighboring text. See [POWERPOINT_SCOPE.md](POWERPOINT_SCOPE.md).

## Word Blocks reflow their SVG frame, not their picture transform

PowerPoint exposes `AfterShapeSizeChange`; Word's COM model does not. A fixed Word Block nevertheless has the same
underlying contract whether it is an `InlineShape` or a floating `Shape`: its TeX measure and exact outer SVG frame are
separate, and reflow rebuilds the root viewport without stretching TeX coordinates. The add-in uses a documented
`EVENT_SYSTEM_CAPTUREEND` WinEvent scoped to the current WINWORD process. Its native callback only posts a one-shot
task to the VSTO UI thread; the UI code reads the final geometry after mouse-up and queues the renderer. It is neither
a geometry poll, document watcher, global low-level mouse hook, nor Office-window subclass. `WindowSelectionChange`
is retained as a fallback for non-mouse changes or unavailable operating-system monitoring. **Reflow Frame** remains
the explicit path. Moving and rotating do not rerender. This restores the PowerPoint semantic boundary—end of a resize
gesture—without persisting a distorted SVG image transform.

## Word equation numbers use tabs and fields, not a table

A numbered display equation stays on a manual-break visual line in the current paragraph. A center tab aligns the
formula; a right tab aligns a native `SEQ LaTeXBlockEq` field. The add-in registers the matching `LaTeXBlockEq`
Caption Label so Word recognizes the category, but Word's native Cross-reference dialog does not treat bare `SEQ`
fields in this tab scaffold as caption objects. Stable bookmarks on the field results therefore remain the
document-portable per-equation target layer, and the add-in inserts native `REF` fields through its own equation
picker. This retains Word search and cross-reference semantics
without adding a table or a paragraph solely for the equation.
The current picker and **Update Numbers** operation deliberately cover the main document story only; headers,
footnotes, comments, and text boxes are outside this first cross-reference scope.

## Formula-as-an-OpenType-glyph remains an experiment

The font-glyph spikes demonstrate that a formula carried by a real Word character receives ordinary character spacing
and can be embedded in a DOCX. They do not provide a product backend: arbitrary TeX would require a deterministic
glyph allocator, font construction/subsetting, embedded-font persistence, correct PUA behavior, SVG-feature coverage,
and a licensing strategy. The experiments remain under `samples/`; SVG is the supported representation.

## StemTeX is an external dependency, not a submodule

StemTeX has its own source repository and release lifecycle. LaTeX Blocks consumes a staged runtime during development
and freezes a compatible runtime and profiles into each installer. See [STEMTEX_INTEGRATION.md](STEMTEX_INTEGRATION.md).

## No broad document watchers

Preview rendering is asynchronous, but document changes occur only through explicit user commands or narrowly scoped
host events. The project avoids document-wide polling, continuous field renumbering, and background text-space
rewrites. This is necessary for predictable Office responsiveness and document ownership.
