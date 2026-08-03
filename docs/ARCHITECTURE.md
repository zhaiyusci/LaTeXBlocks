# Architecture

LaTeX Blocks is an Office integration layer, not a second TeX engine. StemTeX owns TeX layout and SVG rendering;
Word and PowerPoint own document/slide placement, host UI, and persistence. This separation is the central design
constraint.

## System boundary

```text
Office Ribbon / editor
        │
        ▼
Host-specific service ──────► StemTeXBackend ──────► StemTeX native runtime
        │                           │                         │
        ▼                           │                         ▼
Word InlineShape or                │                    SVG + TeX metrics
PowerPoint SVG shape ◄─────────────┘
        │
        ▼
Office document or presentation
```

The add-ins embed the rendered SVG bytes directly. Existing documents therefore retain their visual content even
when StemTeX is unavailable; StemTeX is needed only to insert or rerender a block.

## Host responsibilities

| Concern | Word | PowerPoint |
| --- | --- | --- |
| Visual object | `InlineShape` SVG | Positioned SVG shape |
| Text integration | Inline formulas, fixed blocks, and numbered equation lines | Free-standing blocks only |
| Source | `InlineShape.AlternativeText` | Shape Alternative Text |
| Identity | Compact metadata in `Title` | Metadata in `Title` plus a dedicated shape tag |
| Layout | Baseline, paragraph, tabs, and document line layout | Position plus one native host frame; every native size change queues TeX reflow |

The shared rendering and metadata code does not imply shared host layout code. Word-specific baseline, U+2060,
paragraph, and tab-stop behavior must never leak into PowerPoint.

### PowerPoint host-frame contract

Every PowerPoint LaTeX Block is one native host frame around unscaled TeX SVG content. All PowerPoint resize handles
use the same contract; there is no side-handle, corner-handle, or vertical-handle zoom mode. A width or height change
is debounced and sent through the asynchronous StemTeX layout path. A changed width derives a fresh typesetting
measure from the prior real SVG root; a height-only change still rerenders the current measure. Translation and
rotation do not enter that path. The TeX coordinates, aspect ratio, and physical scale remain 1:1; if no fixed-size
TeX layout can meet a constrained dimension, the frame grows to the natural safe extent rather than stretching or
cropping content.

`VisualScale` is deliberately absent from the model. The persisted native shape geometry is the host frame, while
the rendering metadata records the TeX layout width and design size that produced its content.

## Rendering lifecycle

Each Office host owns one `StemTeXBackend`, one current profile, and one dedicated FIFO background worker. Native
renderer creation, rendering, disposal, and profile changes occur off the Office UI thread.

- **Preview work** is latest-only. The editor debounces source changes, skips superseded queued work, and discards a
  result if its UI request ID is stale.
- **Document mutations** are durable. A completed insert, edit, explicit typesetting-width change, font-size refresh,
  or host-frame update is queued so that a later preview cannot silently cancel it.
- **Office shutdown** invalidates managed work and returns without joining a native-renderer thread on the Office UI
  thread. A background reaper handles only worker processes owned by that host.

This mirrors the responsiveness model of StemTeX GUI while keeping all Word/PowerPoint COM mutations on the
appropriate host UI path.

## Object persistence

The visual SVG and its source are one semantic object, not an image plus a duplicate hidden-text record.

- The SVG is portable display output.
- Alternative Text is the authoritative TeX source, normalized to LF line endings.
- Title metadata stores only identity and rendering/layout facts needed to edit the object.
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
