#!/usr/bin/env python3
"""Tell Word to use Formula Glyph Character Spike in every font slot for U+E000.

Word's COM API writes the PUA character with ``w:hint="eastAsia"`` but its
``Font.NameFarEast`` setter is unavailable in this Office installation.  This
small OOXML post-step is deliberately explicit: it makes the same run use the
spike font for ``ascii``, ``hAnsi``, ``eastAsia``, and ``cs``. It does not
create any drawing or alter text content.
"""

from __future__ import annotations

import argparse
import os
import shutil
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


FONT_NAME = "Formula Glyph Character Spike"
PUA = "\ue000"


def patch_document(document: Path) -> None:
    with ZipFile(document, "r") as source:
        xml = source.read("word/document.xml").decode("utf-8")
        before, marker, after = xml.partition(PUA)
        if not marker:
            raise RuntimeError("U+E000 was not found in word/document.xml.")

        # The immediately preceding run-properties block is the glyph's run.
        run_start = before.rfind("<w:r>")
        rpr_end = before.rfind("</w:rPr>")
        if run_start < 0 or rpr_end < run_start:
            raise RuntimeError("Could not locate the PUA run properties.")
        run_prefix = before[run_start:rpr_end]
        fonts_start = run_prefix.rfind("<w:rFonts")
        fonts_end = run_prefix.find("/>", fonts_start)
        if fonts_start < 0 or fonts_end < fonts_start:
            raise RuntimeError("Could not locate w:rFonts for U+E000.")

        fonts = run_prefix[fonts_start : fonts_end + 2]
        import re

        # The COM-generated run always contains ascii, hAnsi, and cs; use a
        # generic helper so reruns are deterministic too.
        def set_font_slot(value: str, slot: str) -> str:
            attribute = f'w:{slot}'
            if f'{attribute}=' in value:
                return re.sub(
                    rf'{re.escape(attribute)}="[^"]*"',
                    f'{attribute}="{FONT_NAME}"',
                    value,
                )
            return value[:-2] + f' {attribute}="{FONT_NAME}"/>'

        for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts = set_font_slot(fonts, slot)

        # Word's COM layer marks a PUA character with hint="eastAsia" even
        # when every rFonts slot names this font.  On this Office build that
        # hint causes a blank fallback lookup for U+E000.  With all four slots
        # explicit, removing the hint makes Word take the ordinary hAnsi path.
        fonts = re.sub(r'\s+w:hint="[^"]*"', "", fonts)

        patched_before = before[:run_start] + run_prefix[:fonts_start] + fonts + run_prefix[fonts_end + 2 :] + before[rpr_end:]
        patched_xml = patched_before + marker + after

        fd, temp_name = tempfile.mkstemp(suffix=".docx", dir=document.parent)
        os.close(fd)
        Path(temp_name).unlink(missing_ok=True)
        temp = Path(temp_name)
        try:
            with ZipFile(temp, "w", ZIP_DEFLATED) as destination:
                for entry in source.infolist():
                    data = patched_xml.encode("utf-8") if entry.filename == "word/document.xml" else source.read(entry.filename)
                    destination.writestr(entry, data)
            shutil.move(temp, document)
        finally:
            temp.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("document", type=Path)
    args = parser.parse_args()
    patch_document(args.document)
    print(f"Patched the East-Asian font slot in {args.document}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
