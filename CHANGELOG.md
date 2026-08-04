# Changelog

All notable product-facing changes are recorded here. Version history begins with the first self-contained
Word-and-PowerPoint package line.

## [Unreleased]

## [0.2.26] — 2026-08-04

### Word

- Fixed a color-rendering regression where the wrapper could append a TeX word space to an auto-width inline formula.
  Word Font Color now changes paint only: the exact TeX box width is unchanged, including when the source ends in a line break.

## [0.2.25] — 2026-08-04

### Word

- Inline formulas, fixed-width blocks, and numbered display equations now inherit the native Word **Font Color** at
  insertion. Recoloring an existing block uses that same Word formatting as the source of truth and asynchronously
  rerenders its SVG without changing the authoritative LaTeX stored in Alternative Text.

## [0.2.24] — 2026-08-04

### PowerPoint

- Fixed fixed-height bare blocks with **Vertical: Top**: the TeX SVG viewport now
  begins at the host frame's top edge instead of remaining vertically centered.
  Middle and Bottom retain their respective viewport placements, including when
  the block otherwise uses the default style.

## [0.2.23] — 2026-08-04

### PowerPoint

- Native PowerPoint frame dimensions are now authoritative. When a reflowed TeX block still exceeds a user-shrunk
  frame, its SVG keeps the exact requested viewport and clips overflow instead of expanding back to a natural size.

## [0.2.22] — 2026-08-04

### PowerPoint

- Bundled a corrected StemTeX worker template for full-width request content: the worker now suppresses its own outer
  paragraph indentation before starting the request minipage.
- Moved PowerPoint block padding, background, border, and vertical placement out of TeX boxes and into the final SVG
  shell. Typography (leading and text color) remains in TeX. SVG borders are four in-viewport filled strips, so the
  trailing edge cannot be clipped by a centered stroke or an incorrect TeX box viewport.
- Added a regression check that verifies every generated frame rectangle fits inside the emitted SVG `viewBox`.

## [0.2.16] — 2026-08-04

### Reliability

- PowerPoint now defers the embedded preview browser until the Ribbon callback has returned and retries transient OLE
  `RPC_E_SERVERCALL_RETRYLATER` / call-rejected responses. A temporarily busy PowerPoint instance no longer reports a
  successful TeX render as a preview failure.

## [0.2.15] — 2026-08-04

### PowerPoint

- Unified every native PowerPoint resize handle under one host-frame contract. Every actual size change now queues a
  real asynchronous StemTeX layout pass: width changes derive a new stored typesetting measure, while height-only
  changes rerender the current measure. Translation and rotation do not reflow. The TeX SVG remains 1:1 and is never
  stretched or cropped.
- Removed the `VisualScale` concept entirely. Actual formula size is controlled only by **TeX size (pt)** and always
  rerenders.
- Added per-block styling in the PowerPoint editor: ordinary-paragraph line spacing, uniform padding,
  Top/Middle/Bottom placement, text color, background fill, and border color/width. The original author source
  remains in Alternative Text.

## [0.2.14] — 2026-08-03

### Reliability

- Preview cancellation is now isolated from queued insert/update renders. Closing an editor promptly cancels only
  obsolete preview work, and the renderer can recover from a failed or canceled profile initialization.
- Word and PowerPoint lifecycle, profile-switch, document mutation, undo, and Office-event paths now clean up
  transactionally. A failed operation preserves the existing object instead of partially replacing it.
- PowerPoint block recovery preserves geometry and visual-scale metadata during exceptional render/update paths.

### Verification and packaging

- Release smoke coverage exercises active-preview cancellation, renderer recovery, immediate shutdown, U+2060
  persistence, Word equation numbering, and PowerPoint replacement geometry.
- The release procedure now validates the installed PowerPoint VSTO package. VSTO cannot safely swap a solution
  identity between an installed codebase and a development directory through registry edits alone.

## [0.2.13] — 2026-08-03

### Word

- Auto-width inline formulas now use an exact TeX SVG box with a U+2060 WORD JOINER immediately on each side.
  Existing boundaries are reused on edit, adjacent formulas share their middle boundary, and a conversion to a fixed
  block removes unshared boundaries.
- The old horizontal signed-`wp:effectExtent` / adjacent-space measurement path is removed. All drawing effect
  extents are normalized to zero; TeX depth remains the only baseline-mapping input.
- Word smoke coverage now verifies U+2060 insertion, repeated updates, save/reopen, caret placement, adjacent
  formulas, and Auto-to-Fixed conversion against desktop Word.

### Packaging

- The `0.2.13` installer publishes both VSTO add-ins and bundles StemTeX `0.12.4` with the supported profiles.
- Documentation now separates product scope, object contracts, StemTeX integration, developer workflows, testing,
  release operations, and design decisions.
