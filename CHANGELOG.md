# Changelog

All notable product-facing changes are recorded here. Version history begins with the first self-contained
Word-and-PowerPoint package line.

## [Unreleased]

## [0.2.38] — 2026-08-05

### Block vertical alignment

- Styled Word and PowerPoint Blocks now align their Top/Middle/Bottom frame positions
  against a real TeX text line box rather than the visible ink bounds of the first
  glyph. A lowercase `x`, capital `A`, and descender-bearing `g` therefore reserve
  the same ascender/depth space at a given font size and selected leading.
- The line-box strut is rebuilt after the final leading is selected, so its height
  follows the editor's line-spacing value instead of a stale profile default.
- Standalone display math (`\[...\]`, equation/align-style environments) preserves
  its own TeX vertical list exactly: no wrapper paragraph, strut, or outer TeX
  colour command is injected. Its default foreground colour is inherited from the
  SVG root, while explicit colours in the author source still take precedence.

## [0.2.37] — 2026-08-05

### Word

- Fixed Content Blocks now have the same persistent style editor as PowerPoint: TeX font size,
  line spacing, uniform padding, Top/Middle/Bottom placement, text color, background fill, and
  border color/width.
- The shared style model keeps typography in TeX and the outer shell in SVG. Styled Word Blocks
  preserve their source in Alternative Text, persist their style in compact Title metadata, and
  repaint padding/fill/border at the correct edges after an inline or floating frame resize.
- Word and PowerPoint now give an editor-confirmed default style its literal meaning: 1.20× leading
  is authored in TeX and the SVG owns the viewport. Existing default blocks stay on their compatible
  bare-SVG route until edited, so opening old documents does not reformat them.
- Inline formulas and Word-native numbered equations remain deliberately unstyled so their
  running-text baseline and same-paragraph tab/field semantics are unchanged.

## [0.2.36] — 2026-08-05

### Word

- Hardened fixed-Block frame reflow around real Word gestures: the add-in now compares the
  frame before and after each gesture, so moving or rotating a Block never queues a render.
  Rapid consecutive resizes accumulate from the latest intended TeX measure rather than from
  stale document metadata.
- Reflow work is now keyed to the actual Word drawing object, rather than copied metadata, so
  independently resized copies cannot overwrite one another. A native text-colour or width
  refresh for a fixed Block is committed through the same framed SVG path.
- Physical Block frames no longer silently cap at 2000 pt. The TeX layout-width policy remains
  bounded at 30–2000 pt, while a valid user-owned SVG frame is preserved exactly.
- Fixed Content Blocks use this same resize-on-release contract in line with text and under every
  ordinary floating wrapping mode. Flow participation, moving, rotating, and changing wrapping
  never themselves cause a re-render.

### Rendering lifecycle

- Moved the native StemTeX renderer behind the x64 **LaTeXBlocks Render Host** process. Word and
  PowerPoint now use a versioned local named-pipe protocol for profile warm-up, latest-only
  previews, durable renders, and cancellation; Office never owns an in-process native renderer.
- The Office add-in owns its Render Host through a Windows Job Object. VSTO unload and application
  shutdown release the job immediately, so Word does not wait for a native renderer create or
  render call to return.

## [0.2.33] — 2026-08-05

### Word

- Fixed-width ordinary LaTeX Blocks now use one resize/reflow contract whether Word keeps them in line with text or
  exposes them as a floating object under any wrapping mode. Changing either outer-frame axis rerenders the TeX box
  and rebuilds an exact SVG viewport rather than persisting a Word image-scale transform.
- Native resize commits now begin at Word's mouse-capture end (mouse-up), via a process-scoped WinEvent whose callback
  performs no COM work and schedules one UI-thread continuation. The existing selection-transition path remains a
  non-polling fallback; **Reflow Frame** accepts either fixed Block representation.
- Editing a resized inline fixed Block now preserves its author-owned outer frame, matching floating Block editing.

## [0.2.32] — 2026-08-05

### Word

- Floating fixed-width LaTeX Blocks now persist an exact SVG frame separately from their TeX layout width. Native
  frame changes are rerendered and reframed without stretching TeX artwork; a width change derives a fresh measure,
  while a height-only change preserves the measure and clips or adds viewport space as needed.
- Added **Reflow Frame** for the selected floating Block. Word has no shape-resize event, so the same operation is
  also queued asynchronously when the selection leaves a resized Block; moving or rotating it does not rerender.

## [0.2.30] — 2026-08-05

### Word

- Fixed-width LaTeX Blocks remain editable after Word Layout Options converts them to floating SVG objects. Editing
  preserves their floating position, relative frame, wrapping, margins, supported object formatting, and metadata.

## [0.2.29] — 2026-08-05

### Word

- Removed the unsuccessful post-command Font Size refresh experiment. Existing selection-change refresh remains the
  supported native-format synchronization path.

## [0.2.27] — 2026-08-04

### Word

- Numbered equations now have an **Equation Reference** picker. It inserts a native, hyperlink-enabled Word `REF`
  field to the individual equation number, so references persist in DOCX files and follow **Update Numbers**.

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
