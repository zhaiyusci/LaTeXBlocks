# Changelog

All notable product-facing changes are recorded here. Version history begins with the first self-contained
Word-and-PowerPoint package line.

## [Unreleased]

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
