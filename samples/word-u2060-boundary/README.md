# U+2060 inline-image boundary experiment

`Build-U2060BoundaryExperiment.ps1` creates an actual Microsoft Word document, not an add-in simulation. It is the
layout experiment behind the production Word Auto-formula placement scaffold. It compares a raw inline SVG image with
the same image enclosed by U+2060 WORD JOINER characters:

```text
A U+0020 [InlineShape] U+0020 B
A U+0020 U+2060 [InlineShape] U+2060 U+0020 B
```

The inline shape has `AlternativeText = "LaTeX source: $E = mc^2$"`; the test does not use `wp:effectExtent` to
change spacing. The result demonstrates that Word's direct image-adjacency path is host-font dependent: U+2060
restores the ordinary space advance in the Times New Roman case, while the SimSun result differs.

The production add-in uses this boundary placement only for auto-width ordinary-content formulas. It treats the SVG
as the exact TeX box, adds no horizontal `effectExtent` or negative TeX space, and reuses existing immediate joiners
on update so they do not accumulate. The add-in's Word smoke test—not this standalone sample—covers repeated edits,
save/reopen, adjacent formulas sharing a joiner, caret placement, and Auto-to-Fixed conversion. See
[`docs/OBJECT_MODEL.md`](../../docs/OBJECT_MODEL.md) for the normative contract.
