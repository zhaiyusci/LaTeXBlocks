# Architecture

LaTeX Blocks is an Office integration layer, not a second TeX engine. StemTeX owns TeX layout and SVG rendering;
Word and PowerPoint own document/slide placement, host UI, and persistence. This separation is the central design
constraint.

## System boundary

```text
Office Ribbon / editor
        │
        ▼
Host-specific service ──────► RenderHost client ────── named pipe ──────► LaTeXBlocks.RenderHost.host (x64)
        │                           │                                      │
        ▼                           │                                      ▼
Word SVG object or                 │                              StemTeX native runtime
PowerPoint SVG shape ◄─────────────┴──────────────────────────── SVG + TeX metrics
        │
        ▼
Office document or presentation
```

The add-ins embed the rendered SVG bytes directly. Existing documents therefore retain their visual content even
when StemTeX is unavailable; StemTeX is needed only to insert or rerender a block.

`LaTeXBlocks.RenderHost.host` is a per-Office-host, x64 child process. The VSTO add-in contains only the
host-specific UI/document code and a named-pipe client; it does not load the native renderer or wait for native TeX
work during Office teardown. The client creates a unique current-user pipe name, launches the render host, and
places it in a kill-on-close Windows Job Object. Releasing the add-in closes that job handle, which terminates the
render host and any processes it owns without making Word or PowerPoint wait for a render, warm-up, or cleanup.

## Host responsibilities

| Concern | Word | PowerPoint |
| --- | --- | --- |
| Visual object | SVG `InlineShape`, or a floating SVG `Shape` for a fixed Content Block | Positioned SVG shape |
| Text integration | Inline formulas, fixed blocks, and numbered equation lines | Free-standing blocks only |
| Source and identity | Magic header plus TeX source in SVG Alternative Text (`InlineShape` or `Shape`) | The same magic-header envelope in Shape Alternative Text |
| Block decoration | Fixed Content Block style in the magic header: TeX renders typography; the SVG layer renders the outer shell | The same declared style fields in the magic header |
| Layout | Baseline, paragraph, tabs, document line layout, and the fixed Content Block frame contract below | Position plus one native host frame; every native size change queues TeX reflow |

The shared rendering and metadata code does not imply shared host layout code. Word-specific baseline, U+2060,
paragraph, and tab-stop behavior must never leak into PowerPoint.

### PowerPoint host-frame contract

Every PowerPoint LaTeX Block is one native host frame around unscaled TeX SVG content. All PowerPoint resize handles
use the same contract; there is no side-handle, corner-handle, or vertical-handle zoom mode. A width or height change
is debounced and sent through the asynchronous StemTeX layout path. A changed width derives a fresh typesetting
measure from the prior real SVG root; a height-only change still rerenders the current measure. Translation and
rotation do not enter that path. The TeX coordinates, aspect ratio, and physical scale remain 1:1. A native target
frame is authoritative: if reflow cannot make its TeX content fit, the final SVG uses that exact viewport and clips
the overflow rather than stretching the artwork or growing the frame.

`VisualScale` is deliberately absent from the model. The persisted native shape geometry is the host frame, while
the rendering metadata records the TeX layout width and design size that produced its content.

### Word fixed-Content frame contract

For Word, **fixed Content Block** is the semantic category `role=Content` plus `mode=Fixed`; it is not a synonym
for a floating object. The same external-frame contract applies when Word exposes that Block as an `InlineShape`
(**In Line with Text**) and when the user changes it into a floating `Shape` with any normal floating wrap mode:
**Square**, **Tight**, **Through**, **Top and Bottom**, **Behind Text**, or **In Front of Text**.

A native width or height change is treated as an instruction to re-typeset and then rebuild the exact physical SVG
viewport after the resize gesture ends. Moving, rotating, or merely changing the wrapping mode does not re-typeset.
The viewport preserves TeX's physical coordinate scale: enlarging it adds space and shrinking it clips; neither
operation persistently stretches mathematical glyphs. The TeX viewport is left-anchored, so added width belongs on
the right and a narrow frame clips the right edge. Auto-width inline formulas and numbered equations are
different layout roles and do not acquire this independent frame behavior.

## Rendering lifecycle

Each Office host owns one `RenderHostClientBackend`, one current profile preference, and one disposable external
`LaTeXBlocks.RenderHost.host` process. The RenderHost—not Word or PowerPoint—owns the native StemTeX lifetime and
its worker queue. Office COM mutations remain on the appropriate host UI path.

- **Startup warm-up** is nonblocking. The add-in queues a profile switch over the pipe and can immediately return to
  the Office UI; a later render waits in the RenderHost queue if warm-up is still underway.
- **Preview work** is latest-only. The editor debounces source changes; a newer preview or explicit cancellation is
  sent over an independent pipe connection and supersedes the older preview. Stale results are discarded by request
  ID.
- **Document mutations** use durable queued render requests. An insert, edit, explicit typesetting-width change,
  font-size refresh, or host-frame update is not silently cancelled by a later preview.
- **Office shutdown** closes the RenderHost Job Object and returns without joining, reaping, or unloading native
  renderer work inside the Office process.

This mirrors the responsiveness model of StemTeX GUI while keeping all Word/PowerPoint COM mutations on the
appropriate host UI path.

## Object persistence

The visual SVG and its source are one semantic object, not an image plus a duplicate hidden-text record.

- The SVG is portable display output.
- Alternative Text contains the versioned magic header followed by the authoritative TeX source; see [MAGIC_HEADER.md](MAGIC_HEADER.md).
- Title is empty after a committed operation. No Title/JSON or shape-tag compatibility format is read.
- For Word inline and numbered formulas, the drawing run's native `Font.Color` remains the authoritative text color.
  Fixed Content Blocks instead persist a declarative block style in their magic header; their editor is the
  authoritative color/leading/padding/vertical-placement UI. Neither route copies visual settings into Alternative
  Text.
- Both hosts share one style model. Before a styled fixed-block preview or committed render, the service subtracts
  the configured padding on all four sides from the authored outer frame and constructs a scoped TeX box with that
  exact content width and height. TeX sets paragraph indentation to zero, keeps horizontal layout left-aligned, and
  owns leading, text color, and Top/Middle/Bottom placement inside its fixed-height `vbox`. Its ordinary-text branch
  gives the outer text lines a stable typographic height/depth so lowercase-only content cannot collapse to x-height;
  standalone display math receives no paragraph or line-box injection. The SVG root places that returned box at the padding origin, paints the background and inside border, and
  clips the authored outer frame. It performs no second vertical-placement calculation, never rewrites Alternative
  source portion, and does not rely on Office Fill/Line formatting for a
  block's visible decoration. An explicit style is meaningful even at its apparent defaults: 1.20× leading is authored
  in TeX.
- A successful edit creates and annotates a replacement SVG before removing the old visual object.

The detailed Word representation is normative in [OBJECT_MODEL.md](OBJECT_MODEL.md). PowerPoint's deliberately
narrower object model is specified in [POWERPOINT_SCOPE.md](POWERPOINT_SCOPE.md).

## Profile and runtime boundary

Profiles belong to a host session preference rather than to a document object: Word and PowerPoint persist separate
choices. The renderer discovers usable profiles from the selected StemTeX installation. Runtime discovery,
compatibility, and packaging are specified in [STEMTEX_INTEGRATION.md](STEMTEX_INTEGRATION.md).

## Source layout

```text
src/LaTeXBlocks.Word.AddIn/        Word VSTO entry point, Word service, and shared renderer bridge
src/LaTeXBlocks.PowerPoint.AddIn/  PowerPoint VSTO entry point and slide-object service
tests/                             Real Office smoke tests
scripts/                           Explicit build, registration, test, and publishing workflows
installer/                         Self-contained per-user installer
```

The PowerPoint project links the shared renderer, metadata, and editor code where appropriate, but its service
remains separate so host behavior stays explicit and testable.
