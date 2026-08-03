#!/usr/bin/env python3
"""Build a deliberately small, reproducible PUA-font experiment for Word.

This is not a production math-font generator.  It composes the outlines for
``E = mc²`` from DejaVu Sans into one U+E000 glyph, with an ink-tight advance.
The point of the spike is to let Word lay out the formula as a *real character*
rather than as an InlineShape.

The generated font contains only .notdef and U+E000.  It is intended only for
the document sample in this directory and uses the installed DejaVu Sans font
as a legal/open source outline source.  See README.md for its licence and the
important limitations of this experiment.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Iterable

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.recordingPen import RecordingPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont


ROOT = Path(__file__).resolve().parent
DEFAULT_SOURCE = Path(r"C:\Windows\Fonts\DejaVuSans.ttf")
DEFAULT_OUTPUT = ROOT / "FormulaGlyphSpike.ttf"
METRICS_OUTPUT = ROOT / "formula-glyph-metrics.json"

PUA_CODEPOINT = 0xE000
ASCII_CONTROL_CODEPOINT = ord("X")
FAMILY_NAME = "Formula Glyph Character Spike"
FULL_NAME = "Formula Glyph Character Spike Regular"
POSTSCRIPT_NAME = "FormulaGlyphCharacterSpike-Regular"


def glyph_name_for_character(font: TTFont, character: str) -> str:
    glyph_name = font.getBestCmap().get(ord(character))
    if glyph_name is None:
        raise ValueError(f"Source font has no glyph for {character!r}.")
    return glyph_name


def record_formula(source: TTFont) -> tuple[RecordingPen, dict[str, float]]:
    """Return vector commands for an intentionally simple E = mc² formula."""

    glyph_set = source.getGlyphSet()
    hmtx = source["hmtx"].metrics
    recorded = RecordingPen()
    x = 0.0

    # (character, scale, vertical offset, extra space before, extra space after)
    # Values are in source-font units.  The superscript is consciously a simple
    # geometric construction; it is not TeX-quality typesetting.
    recipe: Iterable[tuple[str, float, float, float, float]] = (
        ("E", 1.00, 0.0, 0.0, 250.0),
        ("=", 1.00, 0.0, 0.0, 250.0),
        ("m", 1.00, 0.0, 0.0, 0.0),
        ("c", 1.00, 0.0, 0.0, 0.0),
        ("2", 0.62, 720.0, 55.0, 0.0),
    )

    for character, scale, y_offset, before, after in recipe:
        x += before
        glyph_name = glyph_name_for_character(source, character)
        transformed = TransformPen(recorded, (scale, 0, 0, scale, x, y_offset))
        glyph_set[glyph_name].draw(transformed)
        advance, _left_side_bearing = hmtx[glyph_name]
        x += advance * scale + after

    bounds_pen = BoundsPen(None)
    recorded.replay(bounds_pen)
    if bounds_pen.bounds is None:
        raise RuntimeError("The constructed formula has no outlines.")
    x_min, y_min, x_max, y_max = bounds_pen.bounds
    return recorded, {
        "inkLeft": float(x_min),
        "inkBottom": float(y_min),
        "inkRight": float(x_max),
        "inkTop": float(y_max),
        "recipeAdvance": float(x),
    }


def make_notdef() -> object:
    """A small visible .notdef box makes accidental fallback obvious."""

    pen = TTGlyphPen(None)
    pen.moveTo((120, -120))
    pen.lineTo((120, 1580))
    pen.lineTo((900, 1580))
    pen.lineTo((900, -120))
    pen.closePath()
    pen.moveTo((200, -40))
    pen.lineTo((820, -40))
    pen.lineTo((820, 1500))
    pen.lineTo((200, 1500))
    pen.closePath()
    return pen.glyph()


def build_font(source_path: Path, output_path: Path) -> dict[str, object]:
    source = TTFont(source_path)
    units_per_em = source["head"].unitsPerEm
    recorded, bounds = record_formula(source)

    # Shift to an ink-tight glyph: PUA E000's advance is exactly the formula's
    # ink width.  Therefore this spike cannot accidentally prove its case by
    # hiding formula-side padding in normal font side bearings.
    formula_pen = TTGlyphPen(None)
    recorded.replay(
        TransformPen(formula_pen, (1, 0, 0, 1, -bounds["inkLeft"], 0))
    )
    formula_glyph = formula_pen.glyph()
    formula_advance = int(round(bounds["inkRight"] - bounds["inkLeft"]))

    source_os2 = source["OS/2"]
    source_hhea = source["hhea"]
    ascent = max(int(source_hhea.ascent), int(bounds["inkTop"]))
    descent = min(int(source_hhea.descent), int(bounds["inkBottom"]))
    win_ascent = max(int(source_os2.usWinAscent), int(bounds["inkTop"]))
    win_descent = max(int(source_os2.usWinDescent), int(-bounds["inkBottom"]))

    glyph_order = [".notdef", "formulaEmcTwo"]
    fb = FontBuilder(units_per_em, isTTF=True)
    fb.setupGlyphOrder(glyph_order)
    # U+0058 is deliberately an *instrumentation mapping*: it is the same
    # glyph as U+E000 and lets the Word sample distinguish “Word rejected this
    # font” from “Word uses a special path for the PUA code point”.  Production
    # code would expose only allocated PUA code points.
    fb.setupCharacterMap(
        {
            PUA_CODEPOINT: "formulaEmcTwo",
            ASCII_CONTROL_CODEPOINT: "formulaEmcTwo",
        }
    )
    fb.setupGlyf({".notdef": make_notdef(), "formulaEmcTwo": formula_glyph})
    fb.setupHorizontalMetrics(
        {
            ".notdef": (units_per_em // 2, 0),
            "formulaEmcTwo": (formula_advance, 0),
        }
    )
    fb.setupHorizontalHeader(ascent=ascent, descent=descent, lineGap=0)
    fb.setupNameTable(
        {
            "familyName": FAMILY_NAME,
            "styleName": "Regular",
            "uniqueFontIdentifier": f"{FULL_NAME}; generated font spike",
            "fullName": FULL_NAME,
            "psName": POSTSCRIPT_NAME,
            "version": "Version 0.1",
        }
    )
    fb.setupOS2(
        sTypoAscender=ascent,
        sTypoDescender=descent,
        sTypoLineGap=0,
        usWinAscent=win_ascent,
        usWinDescent=win_descent,
        fsType=0,  # installable embedding: Word is permitted to embed it.
    )
    fb.setupPost()
    fb.setupMaxp()
    fb.font["OS/2"].fsSelection = 0x40  # REGULAR
    fb.font["head"].macStyle = 0
    fb.font["name"].setName("Copyright 2026 LaTeX Blocks font-glyph spike", 0, 3, 1, 0x409)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fb.save(output_path)

    metrics: dict[str, object] = {
        "font": output_path.name,
        "family": FAMILY_NAME,
        "postScriptName": POSTSCRIPT_NAME,
        "codepoint": f"U+{PUA_CODEPOINT:04X}",
        "asciiControlCodepoint": f"U+{ASCII_CONTROL_CODEPOINT:04X}",
        "formula": "E = mc²",
        "unitsPerEm": units_per_em,
        "advance": formula_advance,
        "inkBoundsBeforeTightening": bounds,
        "horizontalSideBearings": {"left": 0, "right": 0},
        "embedding": "installable (OS/2 fsType = 0)",
        "sourceOutlineFont": str(source_path),
    }
    METRICS_OUTPUT.write_text(json.dumps(metrics, indent=2) + "\n", encoding="utf-8")
    return metrics


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--source-font",
        type=Path,
        default=Path(os.environ.get("DEJAVU_SANS_TTF", DEFAULT_SOURCE)),
        help="Path to DejaVuSans.ttf (or compatible source outline font).",
    )
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    if not args.source_font.exists():
        raise SystemExit(
            f"Source font was not found: {args.source_font}. "
            "Set DEJAVU_SANS_TTF or pass --source-font."
        )

    metrics = build_font(args.source_font, args.output)
    print(f"Built {args.output}")
    print(json.dumps(metrics, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
