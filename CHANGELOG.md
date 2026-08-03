# Changelog

All notable product-facing changes are recorded here. Version history begins with the first self-contained
Word-and-PowerPoint package line.

## [Unreleased]

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
