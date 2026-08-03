# Formula glyph persistence spike

This directory contains an isolated Word experiment, not production add-in
code.  It tests the only representation that lets a rendered TeX formula take
part in ordinary text shaping: a generated OpenType glyph at U+E000.

`build_formula_glyph_font.py` builds `TeXFormulaGlyphSpike2.ttf`.  Its new
glyph is a composed `E = mc²` formula.  It is mapped both to U+E000 and, only
for the first Word proof, `@`: the ASCII carrier avoids Word's PUA/east-Asian
fallback selection and isolates the question “does a formula-as-glyph receive
ordinary word spacing?”  The associated Word sample inserts that carrier
between ordinary U+0020 spaces.

The intended acceptance checks are:

1. the PUA formula has ordinary text-neighbour spacing;
2. its baseline agrees with surrounding Latin text;
3. a saved document embeds the font and remains readable after the temporary
   font registration is removed.

It deliberately does not claim to be a general TeX-to-font renderer.

After the ASCII-carrier control succeeds, `patch_pua_document.py` creates
`03-pua-formula-glyph.docx`.  It turns just the two custom-font carrier runs
into U+E000 and patches Word's incorrect `w:hint="eastAsia"` choice.  This
tests a real private-use formula character without adding any drawing object
or spacing offset.  The generated font is currently **not** reliably embedded
by Word, so the PUA document is a layout proof on a machine where the test
font is installed—not a portable document proof.

`Set-ExperimentFont.ps1` registers the test font for the current user only,
so a newly launched Word process can see it.  Run it with `-Remove` after the
experiment.  This is intentional: a transient GDI font registration in the
test-process cannot prove that Word, a different process, can use or embed the
font.
