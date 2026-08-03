# Word object comparison

An isolated Word COM experiment for the question: can changing SVG into PNG or
EMF make an inline formula behave like an ordinary character?

It creates `Word-Object-Comparison.docx`, containing the same `E = mc²`
formula represented in four ways:

1. a real Word text run;
2. an SVG `InlineShape`;
3. a PNG `InlineShape`;
4. an EMF `InlineShape`.

Each line keeps the two surrounding literal U+0020 characters visible with a
yellow swatch. The ordinary-text row is the Word baseline reference; the blue
guide inside each graphic marks that asset's declared internal baseline. The
light gray frame shows the graphic object's actual boundary.

Run on Windows with desktop Word installed:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build-WordObjectComparison.ps1
```

The script also exports a Word-produced PDF proof beside the document. Its
`artifacts/` directory is disposable; the DOCX contains embedded copies of the
generated assets.

This is sample-only code. It does not load or modify the LaTeX Blocks add-in.
