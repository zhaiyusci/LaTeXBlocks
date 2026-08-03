# Experiments and samples

This directory contains reproducible Office experiments and visual proof artifacts. They are not product dependencies,
installer inputs, or alternate renderers. The supported architecture is documented in `../docs/`.

| Sample | Question | Project status |
| --- | --- | --- |
| [`word-u2060-boundary`](word-u2060-boundary/README.md) | How does Word space a literal U+0020 next to an inline SVG, and what changes when U+2060 separates it? | Evidence for the production Word Auto-formula boundary scaffold. |
| [`word-object-comparison`](word-object-comparison/README.md) | Does SVG, PNG, or EMF make an inline formula behave like a text character? | Comparative host-behavior evidence; not add-in code. |
| [`font-glyph-spike`](font-glyph-spike/README.md) | Can one formula be carried by an OpenType glyph and shaped as ordinary Word text? | Feasibility spike only. |
| [`font-glyph-persistence`](font-glyph-persistence/README.md) | Can such a formula glyph survive Word font embedding and reopen? | Persistence spike only; not portable product behavior. |

Generated `bin/`, `obj/`, temporary artifacts, and rendered documents are disposable unless a sample README explicitly
calls them out as a retained proof. Conclusions that affect product design belong in
[`../docs/DECISIONS.md`](../docs/DECISIONS.md), not only in an experiment folder.
