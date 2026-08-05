# LaTeX Blocks Render Host

`LaTeXBlocks.RenderHost.host` owns the native StemTeX lifecycle outside Word and
PowerPoint. It contains no Office interop, VSTO, COM selection, or document
logic.

It compiles the existing StemTeX renderer/backend source into this process,
preserving its asynchronous profile warm-up, FIFO durable rendering, and
latest-preview cancellation behavior.

## Pipe protocol

Start it with a unique nonce:

```text
LaTeXBlocks.RenderHost.host --pipe-nonce <16-to-128-character-token>
```

It accepts current-user-only connections on
`LaTeXBlocks.RenderHost.<token>`. Every message is a little-endian `uint32`
length followed by UTF-8 JSON. Requests are capped at 1 MiB. Responses are
capped at 8 MiB because SVG is Base64 encoded; raw SVG is independently capped
at 5 MiB.

Requests have `version: 1`, optional `id`, and `command`. Supported commands:

- `ping`
- `switchProfile` with `profile`
- `renderLatest` with `profile`, `source`, `widthPt`, `autoWidth`, `fontSizePt`
- `renderQueued` with the same render fields
- `cancelLatestPreview`
- `shutdown`

Render successes contain Base64 SVG plus TeX metadata:
`svgBase64`, `depthPt`, `summaryJson`, `outcomeCode`, `issueFlags`, and
`outcomeMessage`.

The server accepts multiple pipe connections and can read later frames while a
render response is pending. Responses can arrive out of order and clients must
match by `id`. Commands are registered in wire order before awaiting their
result, so a following `cancelLatestPreview` properly supersedes the preceding
`renderLatest` instead of racing ahead of it.

`switchProfile` returns immediately after queuing warm-up. `shutdown` sends its
success response before releasing the backend and exiting.

## Build

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe' `
  .\src\LaTeXBlocks.RenderHost\LaTeXBlocks.RenderHost.csproj `
  /p:Configuration=Release /v:minimal
```

The currently bundled StemTeX native SDK is x64, so this host is explicitly
built x64. An x86 StemTeX payload would need a separately built x86 host.
