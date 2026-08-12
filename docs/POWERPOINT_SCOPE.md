# PowerPoint Scope

This document defines the intentionally narrower PowerPoint host. The common rendering architecture is described in
[ARCHITECTURE.md](ARCHITECTURE.md); Word's inline and numbered-equation rules are specified separately in
[OBJECT_MODEL.md](OBJECT_MODEL.md).

## Product boundary

The PowerPoint host supports LaTeX Blocks only. A block is one positioned slide object whose visual content is a
self-contained SVG rendered by StemTeX and whose Alternative Text contains the authoritative LaTeX source.

PowerPoint does not receive an inline-math mode. In particular, the PowerPoint add-in does not:

- insert formulas into a PowerPoint text run;
- simulate a textual baseline or TeX depth against surrounding characters;
- inspect, replace, or compensate neighboring spaces;
- refresh a block continuously in response to PowerPoint text-formatting events; or
- reproduce Word's numbered-equation line and tab-stop scaffold.

These are Word integration concerns, not properties of a LaTeX Block.

Block insertion nevertheless respects the surrounding typography. If the insertion command is invoked with a caret
or selection in ordinary PowerPoint text, the add-in snapshots that text range's point size and sends it to StemTeX as
the block's real TeX design size. A selected text-bearing shape provides the same initial value. Mixed text follows
PowerPoint's insertion convention by using the first/current character size. The editor exposes this size explicitly
and changing it rerenders the TeX; it never rescales an SVG rendered for a different size. When no text context exists,
the editor starts at 18 pt.

## PowerPoint object contract

Each block is a regular positioned PowerPoint SVG shape. The add-in owns only the semantic association between three
parts of that one shape:

| Concern | PowerPoint representation |
| --- | --- |
| Display | Embedded, self-contained SVG |
| Authoritative source | Shape Alternative Text |
| Identification and renderer metadata | Shape title plus a dedicated LaTeX Blocks shape tag |
| Block style | A versioned PowerPoint style tag plus an explicit-style marker |
| Placement and size | Native PowerPoint host-frame geometry; the SVG content remains 1:1 inside that frame |

Selecting a recognized block and invoking **Edit LaTeX Block** reopens its source. A successful render replaces the SVG
atomically while preserving the shape's slide position, rotation, z-order, and stable identity. Ordinary pictures,
MathType objects, and other shapes are never inferred to be LaTeX Blocks merely because they contain SVG or Alternative
Text.

Layout width, host-frame geometry, and TeX design size have separate meanings. A new block starts at StemTeX's
standard 360 pt typesetting width. The editor exposes the same exact point-value control as StemTeX GUI: 30–450 pt,
a 0.5 pt step, and one decimal place. This width is a layout/reflow constraint sent to StemTeX, not an image scale.
That compact control range does not reinterpret an already-saved block or limit direct native frame manipulation.

PowerPoint exposes one host-frame contract for every native resize handle. Any frame-width or frame-height change
starts one debounced asynchronous TeX layout operation; it is never an image-scale operation. A changed width derives
a new stored typesetting width from the prior SVG root. A height-only change rerenders the current stored width, since
StemTeX's block interface has width—not an independent height—as its layout input. A corner drag supplies both frame
constraints in the same operation, not a special zoom mode. Translation and rotation leave the source, TeX layout,
and SVG untouched. **Typesetting width (pt)** remains the direct manual control of the stored layout width, while
**TeX size (pt)** is the only control that changes the TeX design size. Every result uses real TeX font metrics,
scripts, line breaking, and optical-size choices; the add-in never stretches SVG artwork. If a fixed-size TeX layout
still cannot satisfy a user-specified frame after reflow, the root SVG viewport remains exactly that frame and clips
the overflow. The host frame never silently grows back.

The TeX layout area is horizontally left-anchored inside that host frame. Padding alone defines its left inset; the
optional border is painted inside the outer edge without shrinking the content box. Any additional frame width remains
on the right, and narrowing the frame clips the right edge. The SVG shell never recenters the TeX viewport.

The PowerPoint Ribbon deliberately exposes **Insert Block** and **Edit Block**, with no inline-math command.
Double-clicking a recognized block opens the same editor. The Ribbon also exposes selection-aware
**Typesetting width (pt)** and **TeX size (pt)** fields, enabled only for one recognized block. The width field is
the exact stored layout width sent to StemTeX. There is no `VisualScale` property, tag, or compatibility fallback.
Submitting either value rerenders that block asynchronously with the requested layout width or TeX design size. If the user changes slide
or selection before rendering finishes, the replacement remains on the block's owning slide and does not steal the
new selection. Source changes, width changes, size changes, and PowerPoint-profile changes all use the long-lived
asynchronous StemTeX backend; no render is performed synchronously on PowerPoint's UI thread. Live previews are
latest-only, while document changes are merged per block and submitted through a durable FIFO path, so formatting a
second block cannot cancel the first. A failed frame update restores the last valid frame and SVG instead of leaving
stretched artwork. PowerPoint's native `AfterShapeSizeChange` event is used for direct manipulation; there is no
resize polling loop or document-wide geometry watcher.

The selected profile is a PowerPoint-host preference. It is persisted independently from Word, applies to subsequent
PowerPoint preview, insert, and rerender operations, and is not stored on individual blocks.

## Block styling

The PowerPoint editor offers a compact per-block style surface: ordinary-paragraph line spacing, uniform padding,
Top/Middle/Bottom vertical placement, text color, background fill, and border color/width. This is not a facade over
PowerPoint's picture formatting. The add-in keeps the author source verbatim in Alternative Text and serializes only
the style values in a versioned PowerPoint-only tag. At render time, the add-in subtracts padding from the outer
dimensions and StemTeX typesets an exact inner box with zero paragraph indentation, the selected leading and text
color, horizontal left alignment, and the selected Top/Middle/Bottom placement inside a fixed-height `vbox`. No
SVG-side line metric is inferred: ordinary TeX text receives a stable first/final line box, while a standalone display
remains free of paragraph or line-box injection. The add-in places the returned box at the padding origin, then paints
only the background and inside border and clips the finished SVG; it performs no second vertical-alignment calculation. This avoids
treating a TeX `\\fbox` as the outer PowerPoint frame, whose paint can otherwise lie outside
dvisvgm's SVG viewport.

The shared versioned magic header at the start of `AlternativeText` stores block
identity, layout width, font size, and optional style values. The TeX source starts
after the end marker and preserves its own line structure. `Title` remains empty.
Text and background colours are distinct named fields.

New auto-height blocks fit their content, so vertical placement has no spare height to distribute. Once a block has a
fixed native frame height, TeX receives the corresponding inner height and places its content at the selected Top,
Middle, or Bottom position. When content exceeds that explicit frame, the same TeX alignment determines the preserved
edge (or center), and the outer SVG viewport only clips the result; it never scales the TeX output, grows the frame, or
aligns the content again. The line-spacing control applies to ordinary paragraph leading. It deliberately does not redefine
the row spacing of math environments such as `align` or `gather`; those have their own TeX layout controls.

The declared style model is shared with Word's fixed Content Block editor and uses
the same magic-header fields in both hosts. Word inline formulas and numbered
equations deliberately do not use this UI or style shell.

## Shared infrastructure

The PowerPoint and Word add-ins may share the StemTeX backend, profile selection, preview editor, metadata parser, SVG
validation, and atomic-replacement logic. They should not share Word-specific layout code. In particular, PowerPoint
must not inherit Word's `InlineShape`, baseline, adjacent-space, paragraph, tab-stop, or font-size interception paths.
