# Formula glyph as a Word character — spike

This folder tests the only literal interpretation of “make the formula a
character”: put the complete formula in one OpenType glyph and insert its
Private Use Area character (`U+E000`) into Word.

`generate_font.py` uses `fontTools` and the locally installed DejaVu Sans
outline font to construct a deliberately narrow test font:

```text
Formula Glyph Character Spike Regular
    U+E000 = E = mc²
    U+0058 = E = mc²  (diagnostic control only)
```

The glyph has **zero left and right side bearing** and an advance equal to its
real ink width. This is important: the sample cannot appear to solve spacing by
concealing a formula-side margin inside a font.

## Build

```powershell
cd C:\Users\jairy\Documents\LaTeXBlocks\samples\font-glyph-spike
python .\generate_font.py
```

This produces:

- `FormulaGlyphSpike.ttf` — the small experiment font.
- `formula-glyph-metrics.json` — its code point, advance, bounds, and
  embeddability metadata.

The generated output derives the sample outlines from the installed DejaVu
Sans font. It is a test artifact only; any production approach must use a
properly licensed math outline source and retain the applicable licence. The
applicable DejaVu/Bitstream Vera–derived licence is included as
[`LICENSE-DEJAVU`](LICENSE-DEJAVU).

## What this proves, if Word renders it

Word sees `U+E000` as a normal character, not an `InlineShape`. Therefore the
spaces before and after it are normal text spaces. The equation itself gets its
height, depth, and advance from its font glyph metrics.

It does **not** prove that this is ready to be a formula backend:

- This is one hard-coded glyph, not arbitrary LaTeX.
- The superscript is a simple scaled outline, not TeX math layout.
- A full system would need a per-document glyph allocator, an OpenType font
  builder/subsetter, embedded-font persistence, and a mapping from each PUA
  character back to its LaTeX source.
- SVG paint features, colour, clipping, and every special DVI/SVG effect do
  not automatically map to TrueType outlines.

The deliberate `U+0058` (`X`) mapping is only a Word diagnostic: the sample
can show whether a blank result is caused by Word rejecting the generated font
or by a code-point/PUA-selection issue. A real font would not map ordinary
letters this way.

`verify_word.ps1` is the reproducible Word-side test. It adds the generated
font resource, creates a `.docx` and PDF in this folder, then unloads that
transient resource. In this Office build, however, a transient resource is not
enough for a fresh Word process: use `Install-TestFont.ps1`, close/reopen Word,
then run `verify_word.ps1`. `Uninstall-TestFont.ps1` removes only this test
font afterwards.

## Word-specific finding

Word initially writes the PUA run with `w:hint="eastAsia"`. On this machine,
that makes U+E000 fall back to a blank glyph even when every `w:rFonts` slot
names the generated font. `patch_pua_font_slot.py` removes that hint and writes
the family into the `ascii`, `hAnsi`, `eastAsia`, and `cs` slots. The resulting
Word document contains a normal `w:t` character and no `w:drawing` element;
Word’s PDF output visibly renders the PUA formula.

With the current test font, Word writes an `embedRegular` relationship and an
obfuscated `word/fonts/*.odttf` part after `Document.EmbedTrueTypeFonts = true`.
`verify_embedded_font.ps1` can be run after `Uninstall-TestFont.ps1` to confirm
that the saved DOCX still paints from its embedded font. This is promising, but
production still needs deterministic font subsetting, a stable glyph allocator,
and a licensing review.

## Licence note

DejaVu Sans is distributed under a permissive Bitstream Vera–derived licence.
The generated font is not intended for distribution or for use as the product
font; it exists solely to validate Word’s character-layout behavior.
