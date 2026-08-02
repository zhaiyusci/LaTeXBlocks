# PowerPoint Scope

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
| Placement and size | Native PowerPoint shape geometry |

Selecting a recognized block and invoking **Edit LaTeX Block** reopens its source. A successful render replaces the SVG
atomically while preserving the shape's slide position, rotation, z-order, and stable identity. Ordinary pictures,
MathType objects, and other shapes are never inferred to be LaTeX Blocks merely because they contain SVG or Alternative
Text.

Layout width and visual scale are deliberately separate, following Scholia's LaTeX-note model. A new block starts at
StemTeX's standard 360 pt typesetting width. The editor exposes the same exact point-value control as StemTeX GUI:
30–450 pt, a 0.5 pt step, and one decimal place. This is layout/reflow width, not shape scale. That compact control
range does not reinterpret an already-saved block or limit direct horizontal shape manipulation.
After insertion, dragging a left or right side handle changes the StemTeX paragraph width and rerenders once
PowerPoint reports that the resize has finished. A vertical or corner resize changes only the visual scale and reuses
the existing SVG. The TeX design size remains a third, independent property and therefore continues to control font
metrics, scripts, line breaking, and optical-size choices.

The PowerPoint Ribbon deliberately exposes **Insert Block** and **Edit Block**, with no inline-math command.
Double-clicking a recognized block opens the same editor. The Ribbon also exposes selection-aware
**Typesetting width (pt)** and **TeX size (pt)** fields, enabled only for one recognized block. The width field is
the exact stored layout width sent to StemTeX; it never includes independently stored visual scale.
Submitting either value rerenders that block asynchronously with the requested layout width or TeX design size. If the user changes slide
or selection before rendering finishes, the replacement remains on the block's owning slide and does not steal the
new selection. Source changes, width changes, size changes, and global-profile changes all use the long-lived
asynchronous StemTeX backend; no render is performed synchronously on PowerPoint's UI thread. Live previews are
latest-only, while document changes are merged per block and submitted through a durable FIFO path, so formatting a
second block cannot cancel the first. A failed horizontal reflow restores the last valid SVG geometry instead of
leaving stretched artwork. PowerPoint's native `AfterShapeSizeChange` event is used for direct manipulation; there is
no resize polling loop or document-wide geometry watcher.

## Shared infrastructure

The PowerPoint and Word add-ins may share the StemTeX backend, profile selection, preview editor, metadata parser, SVG
validation, and atomic-replacement logic. They should not share Word-specific layout code. In particular, PowerPoint
must not inherit Word's `InlineShape`, baseline, adjacent-space, paragraph, tab-stop, or font-size interception paths.
