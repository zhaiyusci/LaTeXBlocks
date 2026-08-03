"""Turn the ASCII-carrier Word sample into a PUA-carrier Word sample.

Word incorrectly writes ``w:hint="eastAsia"`` for a U+E000 run on this
machine, then ignores its ascii/hAnsi font and falls back to DengXian.  This
post-save patch is deliberately narrow: it patches only custom-font carrier
runs, leaves every other OOXML byte intact, and does not add a drawing, extent,
or spacing workaround.
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path


HERE = Path(__file__).resolve().parent
SOURCE = HERE / "02-font-glyph-embedded.docx"
TARGET = HERE / "03-pua-formula-glyph.docx"
FAMILY = "TeX Formula Glyph Spike 2"


def patch_run(run: str) -> str:
    fonts_match = re.search(r"<w:rFonts\b[^>]*/>", run)
    if not fonts_match or f'w:ascii="{FAMILY}"' not in fonts_match.group(0):
        return run
    fonts = fonts_match.group(0)
    for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
        attribute = f'w:{slot}'
        if attribute + "=" in fonts:
            fonts = re.sub(
                rf'{re.escape(attribute)}="[^"]*"',
                f'{attribute}="{FAMILY}"',
                fonts,
            )
        else:
            fonts = fonts[:-2] + f' {attribute}="{FAMILY}"/>'
    fonts = re.sub(r'\s+w:hint="[^"]*"', "", fonts)
    return run[: fonts_match.start()] + fonts + run[fonts_match.end() :]


def patch_document_xml(xml: str) -> tuple[str, int]:
    changed = 0
    pattern = re.compile(r"<w:r(?:\s[^>]*)?>.*?</w:r>", re.DOTALL)

    def replace(match: re.Match[str]) -> str:
        nonlocal changed
        run = match.group(0)
        if "@" not in run:
            return run
        candidate = patch_run(run)
        if candidate == run:
            return run
        changed += 1
        return candidate.replace("@", "\ue000", 1)

    patched = pattern.sub(replace, xml)
    if changed != 2:
        raise RuntimeError(f"Expected two custom-font carrier runs; patched {changed}.")
    patched = patched.replace(
        "an ASCII carrier from a temporary TrueType font",
        "a U+E000 private-use character from a temporary TrueType font",
    )
    patched = patched.replace(
        "Formula glyph using an ASCII carrier (@)",
        "Formula glyph using U+E000 (Private Use Area)",
    )
    return patched, changed


def main() -> None:
    if not SOURCE.exists():
        raise FileNotFoundError(f"Build the embedded ASCII-carrier sample first: {SOURCE}")
    with zipfile.ZipFile(SOURCE, "r") as source:
        xml = source.read("word/document.xml").decode("utf-8")
        patched_xml, _ = patch_document_xml(xml)
        with tempfile.NamedTemporaryFile(suffix=".docx", dir=HERE, delete=False) as raw_temp:
            temporary = Path(raw_temp.name)
        try:
            with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as destination:
                for entry in source.infolist():
                    data = patched_xml.encode("utf-8") if entry.filename == "word/document.xml" else source.read(entry.filename)
                    destination.writestr(entry, data)
            shutil.move(temporary, TARGET)
        finally:
            temporary.unlink(missing_ok=True)
    print(f"wrote {TARGET}")


if __name__ == "__main__":
    main()
