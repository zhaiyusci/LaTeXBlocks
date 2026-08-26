# Colour ownership model

LaTeXBlocks has two independent colour domains.

## Office graphic colour

`OfficeGraphicColor` is host-owned state. It represents exactly one of:

- Automatic
- direct RGB
- an Office theme slot plus an explicit `None`, `Tint(amount)`, or
  `Shade(amount)` transform

It is applied to the native Office graphic and is never passed to StemTeX as a
reason to render. Word and PowerPoint use separate adapters because their PIA
`ColorFormat` and `FillFormat` interfaces have different COM identities. The
similar API shape does not make those interfaces interchangeable.

The shared model never stores a raw `TintAndShade` or `Brightness` property.
Those are host projections. Word persists the semantic theme slot and variant on
the `w:r` that owns the drawing (`Font.TextColor.TintAndShade`), while its SVG
graphic fill projects the same variant through `Fill.ForeColor.Brightness` and
stores a resolved paint in the embedded SVG/fallback image. Both projections are
parts of one Word graphic state: the run keeps theme responsiveness and the fill
keeps the current SVG paint correct. The Word adapter must not write the two COM
properties onto the same `ColorFormat`. PowerPoint has its own mapping through
shape-fill `TintAndShade`.

Native Office colour commands remain Office operations. LaTeXBlocks observes a
completed Word Font Color command only to mirror the semantic colour onto the
formula graphic fill. It does not rewrite surrounding text and does not render.
For a theme command, the authoritative variant is read only from the one-character
Flat OPC fragment of a formula drawing run. A parser must select the `w:color`
under the `w:r` containing `w:drawing`; selecting the first colour in a document
or selection can silently borrow an unrelated paragraph or text-run colour.

## TeX colour

`TeXColor` is renderer-owned state. It is either `InheritOffice` or an explicit
RGB value selected in LaTeXBlocks UI. An explicit LaTeXBlocks style is written
into the TeX/SVG result and therefore requires a render. Office's native colour
panels do not change this value.

## Replacement rule

Rendering operations such as a TeX font-size or layout-width change snapshot the
`OfficeGraphicColor` before replacing the graphic and restore the same semantic
value afterward. A resolved RGB sample may be used for visual comparison or a
render cache key, but it must never replace Automatic or theme semantics.

Word replacements spanning paragraphs are performed per shape. A single Flat OPC
envelope is restricted to one paragraph because Word can otherwise recreate
paragraph properties and discard direct tab stops used by numbered equations.
