"""Build a valid full TrueType font with one formula glyph at U+E000.

This is deliberately an isolated representation experiment.  It derives from
the open DejaVu Serif TrueType font installed by TeX Live, preserves its
well-tested OpenType tables, and adds a single composed ``E = mc²`` outline.
That makes the result suitable for testing inside Word as a real glyph rather
than an InlineShape.

Run with the workspace Python plus fontTools on PYTHONPATH, for example:

  $env:PYTHONPATH='C:\\tmp\\latexblocks-fonttools'
  & <workspace-python> build_formula_glyph_font.py
"""

from __future__ import annotations

from pathlib import Path

from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont


HERE = Path(__file__).resolve().parent
OUT = HERE / "TeXFormulaGlyphSpike2.ttf"
REGULAR = Path(r"C:\texlive\2026\texmf-dist\fonts\truetype\public\dejavu\DejaVuSerif.ttf")
ITALIC = Path(r"C:\texlive\2026\texmf-dist\fonts\truetype\public\dejavu\DejaVuSerif-Italic.ttf")

FAMILY = "TeX Formula Glyph Spike 2"
FORMULA_GLYPH = "formula_E_equals_mc2"
PUA = 0xE000
ASCII_SPIKE = ord("@")


def add_scaled_glyph(destination: TTGlyphPen, source: TTFont, character: str,
                     scale: float, x: int, y: int) -> int:
    """Copy a TrueType outline and return its scaled advance."""
    glyph_name = source.getBestCmap()[ord(character)]
    advance, _left_side_bearing = source["hmtx"].metrics[glyph_name]
    source.getGlyphSet()[glyph_name].draw(
        TransformPen(destination, (scale, 0, 0, scale, x, y))
    )
    return round(advance * scale)


def make_formula_glyph(regular: TTFont, italic: TTFont):
    """Compose a TeX-like formula into one true glyph on the font baseline."""
    pen = TTGlyphPen(None)
    x = 0
    upm = regular["head"].unitsPerEm

    # These gaps form part of the glyph's own natural advance.  There is no
    # Word-side effectExtent, Font.Spacing, image padding, or negative space.
    x += add_scaled_glyph(pen, regular, "E", 1.0, x, 0)
    x += round(upm * 0.12)
    x += add_scaled_glyph(pen, regular, "=", 1.0, x, 0)
    x += round(upm * 0.12)
    x += add_scaled_glyph(pen, italic, "m", 1.0, x, 0)
    x += add_scaled_glyph(pen, italic, "c", 1.0, x, 0)
    x += add_scaled_glyph(pen, regular, "2", 0.64, x, round(upm * 0.41))
    return pen.glyph(), x + round(upm * 0.012)


def replace_names(font: TTFont) -> None:
    """Give the derived family a unique identity so Word cannot collide it."""
    names = font["name"]
    names.names = [record for record in names.names if record.nameID not in {1, 2, 3, 4, 5, 6, 16, 17}]
    for platform, encoding, language in ((3, 1, 0x0409), (1, 0, 0)):
        names.setName(FAMILY, 1, platform, encoding, language)
        names.setName("Regular", 2, platform, encoding, language)
        names.setName("LaTeXBlocks " + FAMILY + " 1", 3, platform, encoding, language)
        names.setName(FAMILY + " Regular", 4, platform, encoding, language)
        names.setName("Version 1.000", 5, platform, encoding, language)
        names.setName("TeXFormulaGlyphSpike2-Regular", 6, platform, encoding, language)
        names.setName(FAMILY, 16, platform, encoding, language)
        names.setName("Regular", 17, platform, encoding, language)


def build_font() -> None:
    if not REGULAR.exists() or not ITALIC.exists():
        raise FileNotFoundError("DejaVu source fonts were not found under TeX Live.")

    font = TTFont(REGULAR)
    italic = TTFont(ITALIC)
    formula, advance = make_formula_glyph(font, italic)

    if FORMULA_GLYPH not in font.getGlyphOrder():
        font.setGlyphOrder(font.getGlyphOrder() + [FORMULA_GLYPH])
    font["glyf"].glyphs[FORMULA_GLYPH] = formula
    font["hmtx"].metrics[FORMULA_GLYPH] = (advance, 0)
    font["maxp"].numGlyphs = len(font.getGlyphOrder())

    for cmap in font["cmap"].tables:
        if cmap.isUnicode():
            cmap.cmap[PUA] = FORMULA_GLYPH
            # @ is an intentionally non-PUA control mapping for the first
            # Word proof.  It establishes that Word can use the generated
            # font at all, independently from PUA/east-Asian fallback rules.
            cmap.cmap[ASCII_SPIKE] = FORMULA_GLYPH
    font["OS/2"].fsType = 0  # explicitly embeddable for the controlled spike
    font["OS/2"].usFirstCharIndex = min(font.getBestCmap())
    font["OS/2"].usLastCharIndex = 0xFFFF
    replace_names(font)
    font.save(OUT)

    upm = font["head"].unitsPerEm
    print(f"wrote {OUT}")
    print(f"U+{PUA:04X} advance: {advance} units ({advance / upm:.3f} em)")


if __name__ == "__main__":
    build_font()
