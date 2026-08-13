# LaTeX Blocks Magic Header v1

LaTeX Blocks stores its portable object contract and author source together in
the Office object's `AlternativeText`. `Title` is always empty. Word and
PowerPoint use the same format.

## Envelope

The first logical line is the exact version marker `% !latexblocks 1`. It is
followed by field lines of the form `% key: value` and the exact terminator
`% !end-latexblocks`.

The line break immediately following the terminator belongs to the envelope.
Every subsequent character is author source. Implementations must not trim or
otherwise normalize the source except that Office may persist CR and CRLF as LF.

Only a header beginning at character zero is recognized. Keys consist of lower
case ASCII letters and hyphens. Field order has no meaning. A key may occur only
once. Unknown keys are ignored. Values in v1 are restricted to the scalar
formats specified below, so no general string escaping is defined.

## Common fields

Every header contains:

```text
kind: inline-math | display-math | numbered-math | latex-block
id: an RFC 4122 GUID in D form
mode: auto | fixed
font-size-pt: invariant-culture decimal in [1, 200]
```

`width-pt` is required for `fixed` mode and omitted for `auto` mode. It is an
invariant-culture positive decimal.

Formula source is stored as a math body without outer `$...$`, `$$...$$`, or
`\[...\]` delimiters. `kind` preserves the formula form.

## LaTeXBlock style fields

Only `kind: latex-block` may contain these optional fields:

```text
line-spacing: positive invariant-culture decimal
padding-pt: non-negative invariant-culture decimal
vertical-alignment: top | center | bottom
text-color: #RRGGBB
background-color: #RRGGBB
border-width-pt: non-negative invariant-culture decimal
border-color: #RRGGBB
```

Every `kind: latex-block` header also contains these required fields:

```text
frame-width-pt: positive invariant-culture decimal
frame-height-pt: positive invariant-culture decimal
```

They record the dimensions of the last committed SVG frame. Office resize
events are delivered after the live shape has changed, so this committed frame
is the comparison baseline that distinguishes a user resize from an unchanged
object. A successful refresh replaces both values with the new SVG frame.

An absent `background-color` means no background. An absent or zero
`border-width-pt` means no border. If no style field appears, the current
default style is used. When any style field appears, omitted style fields take
their normal defaults.

The contract does not persist derived runtime state: `role` and TeX depth are
excluded. Numbered role follows from `kind`; TeX depth comes from a render
result. The committed frame dimensions above are durable authoring state, not
the transient dimensions observed from an Office shape during a resize.

## Examples

```tex
% !latexblocks 1
% kind: inline-math
% id: a73349fd-73c1-4fec-9b28-90707115a29a
% mode: auto
% font-size-pt: 12
% !end-latexblocks
E=mc^2
```

```tex
% !latexblocks 1
% kind: latex-block
% id: 4c539acd-ef35-4b44-86a4-779a3dc6c886
% mode: fixed
% width-pt: 360
% font-size-pt: 18
% frame-width-pt: 360
% frame-height-pt: 72
% line-spacing: 1.2
% padding-pt: 0
% vertical-alignment: top
% text-color: #000000
% background-color: #FFFF80
% border-width-pt: 0
% border-color: #000000
% !end-latexblocks
Question: is $\mathbf{K}$ the only choice?

This is a second paragraph.
```

## Host contract

- `AlternativeText` contains the complete envelope.
- `Title` is empty after every committed insert, update, conversion, or refresh.
- Editors, renderers, and export operate on author source, never on the header.
- Comprehensive Find is unchanged and continues to search raw
  `AlternativeText`, including the header.
- An object is a LaTeX Blocks object only when the complete v1 header and all
  required fields parse successfully.
- No legacy Title/JSON or tag format is read.
