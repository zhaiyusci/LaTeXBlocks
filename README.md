# LaTeX Blocks

LaTeX Blocks inserts editable LaTeX content into Microsoft Office as portable SVG objects. StemTeX performs the
rendering; the Office host owns only the visual object's placement and document persistence. The repository contains
separate Word and PowerPoint VSTO add-ins. PowerPoint deliberately has a narrower object model: it inserts only
free-standing LaTeX Blocks and does not offer inline math.

This repository intentionally does not implement document search. [Comprehensive Find](https://github.com/zhaiyusci/ComprehensiveFind) can search ordinary Word text and object Alternative Text as a single result space, so LaTeX source remains discoverable without hidden proxy text.

## Host scope

- **Word** supports inline formulas, fixed-width LaTeX blocks, and Word-native numbered display equations. Inline-only
  behavior such as running-text baseline mapping, adjacent-space compensation, and host font-size refresh belongs to
  the Word add-in.
- **PowerPoint** supports one semantic object: a positioned LaTeX Block. It may be moved and resized as a slide shape,
  selected for source editing, and found through its Alternative Text. Insertion snapshots the current ordinary text
  size and passes that point size to StemTeX even though the block itself remains a free-standing shape. It does not
  participate in a text run and has no inline-formula command, baseline compensation, or neighboring-space logic.

See [docs/POWERPOINT_SCOPE.md](docs/POWERPOINT_SCOPE.md) for the PowerPoint boundary.

## Current implementation

PowerPoint:

- Insert and edit a fixed-width, free-standing LaTeX Block as a native SVG slide shape.
- Initialize the TeX design size from the active ordinary-text caret or selection, or from a selected text-bearing
  shape. The point size is passed to StemTeX as a renderer input; the add-in does not scale an SVG rendered at another
  TeX size.
- Start a new block with StemTeX's standard 360 pt typesetting width. The editor exposes that exact layout width in
  points, independently of the shape's visual scale.
- Reflow a block by dragging its left or right side handle; vertical and corner resizing remains independent visual
  scale and does not invoke TeX.
- Set an exact selected-block typesetting width through **Typesetting width (pt)** on the Ribbon.
- Change a selected block directly through the selection-aware **TeX size (pt)** field on the PowerPoint Ribbon; this
  rerenders through StemTeX instead of scaling the existing SVG.
- Preserve the shape's position, independent visual scale, rotation, z-order, stable ID, source, layout width, and TeX size across editing and
  PPTX save/reopen.
- Recognize only shapes carrying the dedicated LaTeX Blocks tag and valid metadata, so ordinary SVG and MathType
  objects are not claimed by the add-in.

Word:

- Insert a traditional inline formula at its TeX natural width.
- Keep ordinary spaces around an inline formula at their normal visual advance without rewriting those U+0020 characters.
- Write the SVG's physical width and height exactly into DrawingML at 12,700 EMU per point, bypassing Word's initial CSS-pixel import rounding without introducing a DPI model.
- Insert a fixed-width LaTeX block for display, multiline, or paragraph content.
- Start a new fixed block with StemTeX's standard 360 pt typesetting width. Fixed blocks use an exact 30–450 pt
  typesetting-width control; auto-width formulas and numbered equations simply show `Natural`.
- Insert a natural-width, display-style equation on its own visual line with a Word-native `SEQ LaTeXEquation` number.
- Update equation numbers explicitly after moving, inserting, or removing complete equation lines; no document watcher is used.
- Preview through the long-lived StemTeX renderer.
- Refresh the editor preview automatically after a short typing pause, while coalescing stale requests.
- Discover StemTeX profiles and persist one global selection for the Word add-in.
- Start warming the default profile in the background as soon as the Word add-in starts.
- Store the authoritative source in the SVG object's Alternative Text, with Word-stable LF line endings.
- Edit a selected block from the Ribbon or by double-clicking it.
- Replace an existing SVG only after the new render succeeds.
- Preserve a stable block ID across edits and DOCX save/reopen cycles.

The smoke test has verified SVG insertion, exact DrawingML extents, adjacent-space compensation across update and
reopen, canonical-source persistence, metadata persistence, atomic replacement, Word-native equation numbering and
bookmarks, deletion-time renumbering, and DOCX reopen in desktop Word.

## Word object contract

The mathematical artifact is an embedded `InlineShape` SVG. Version 1 distinguishes `mode=auto` from `mode=fixed`
and ordinary content from a numbered equation:

- `AlternativeText`: the authoritative LaTeX source, with line endings canonically stored as LF and no duplicate search vocabulary.
- `Title`: short machine metadata in the form `LaTeXBlocks/1;id=<guid>;width=<pt>;depth=<pt>;mode=<auto|fixed>;size=<pt>;role=<content|numbered-equation>`.
- image bytes: a self-contained SVG rendered by StemTeX.

A numbered equation remains inside the current Word paragraph. Manual line breaks create its visual line; a center
tab at the text-column midpoint places the natural-width SVG and a right tab at the text-column edge places
`( { SEQ LaTeXEquation \\* ARABIC } )`. These display tabs deliberately do not bake in running-text paragraph indents.
The field result is
bookmarked with the SVG's stable ID. Word, not StemTeX, therefore owns line placement, numbering, and future
cross-reference semantics without introducing a table or a paragraph boundary.

See [docs/OBJECT_MODEL.md](docs/OBJECT_MODEL.md) for invariants and update behavior.

## Development

Requirements:

- 64-bit Microsoft Word and/or PowerPoint desktop
- Visual Studio with Office/VSTO development components
- .NET Framework 4.8
- a StemTeX stage containing `stemtex-renderer.dll`, `dvisvgmdaemon.dll`, and the `unicodemath_cjk` profile

`STEMTEX_HOME`, when set, selects one runtime explicitly. Otherwise the add-in examines the private installed runtime before adjacent Scholia/StemTeX development stages and selects the usable runtime with the highest semantic version. For equal versions, the installed runtime wins so a development stage cannot silently replace the packaged backend. The current binding requires StemTeX 0.12's native per-request font-size API.

At Office startup each host add-in discovers valid profile directories and queues creation of the renderer for the globally selected profile (`xits_cjk` when no preference has been saved). The Office UI thread is not blocked by the cold start. Like the StemTeX GUI, an add-in owns one renderer and one dedicated FIFO background thread. Renderer creation, rendering, and profile replacement happen on that thread. A profile switch disposes the old renderer and creates the new one. Office shutdown uses a separate fast path: it invalidates managed work and returns without native cancellation, destruction, process enumeration, or a thread join on the Office UI thread. A background reaper terminates only the StemTeX worker tree owned by that host process, including a worker that appears during the initialization race; the exiting process then reclaims the abandoned native renderer.

The editor uses a 250ms input debounce. Every preview receives a monotonically increasing UI request ID. The dedicated
worker skips queued previews that are no longer latest, and the UI discards stale results. Document-format operations
use a separate durable FIFO path: changes are coalesced per block, but a preview or an edit to another block cannot
silently cancel a committed width or TeX-size change. Syntax errors from automatic preview stay in the status line
rather than opening modal dialogs; the Preview button remains available for an immediate refresh with an explicit
error dialog.

After the latest request passes the ID check, its SVG bytes are embedded directly into a fresh preview document. The editor does not navigate to SVG files, so browser navigation state and file caching cannot leave an older formula on screen. The smoke test runs a real hidden editor message loop and verifies that changing `x_1` to `x_2+y_2` changes the committed preview SVG.

Insert and Update are enabled only when the preview exactly matches the current source, mode, width, and global profile. They embed that already-rendered SVG directly; they never repeat native rendering on Word's UI thread.

```powershell
pwsh.exe -NoProfile -File scripts/Initialize-LaTeXBlocks.ps1
pwsh.exe -NoProfile -File scripts/Build-LaTeXBlocks.ps1 -Configuration Debug
pwsh.exe -NoProfile -File scripts/Register-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Both
```

Ordinary builds, tests, and cleans deliberately do not rewrite Word or PowerPoint's live VSTO registration. Use the
explicit `Register-LaTeXBlocks.ps1` command when switching a host to a development build; pass
`/p:EnableVstoProjectRegistration=true` only when the standard VSTO build-time registration behavior is specifically
needed. This prevents a smoke-test build from replacing an installed add-in with `bin\Debug`, or a clean from
unregistering the installed product. Close the selected Office host before reopening it after explicit registration.
The Ribbon tab is named **LaTeX Blocks**. `-TargetHost` accepts `Word`, `PowerPoint`, or `Both`.

Run both end-to-end smoke tests with:

```powershell
pwsh.exe -NoProfile -File scripts/Test-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Both
```

The test executables are explicitly x64 because the current StemTeX SDK and Office installation are x64. The
PowerPoint smoke test verifies ordinary-text size inheritance with a fractional point size, true StemTeX rendering,
shape replacement geometry, explicit identity tags, and PPTX save/reopen persistence.

With PowerPoint closed, the installed add-in's real resize-event path can be checked separately:

```powershell
pwsh.exe -NoProfile -File tests/Test-PowerPointWidthIntegration.ps1
```

That test horizontally resizes a saved block, waits for the asynchronous TeX reflow, verifies the final intrinsic SVG
geometry, closes its temporary presentation without saving, and leaves the installed registration unchanged.

## Installer

Create the self-contained per-user installer with:

```powershell
pwsh.exe -NoProfile -File .\scripts\Publish-LaTeXBlocks.ps1 -Version 0.2.10
```

The package is written to `dist\release\LaTeXBlocks-Setup-<version>.exe`, together with its SHA-256 checksum. It
contains signed Word and PowerPoint VSTO publications, the Microsoft prerequisite bootstrapper, and the StemTeX 0.12
runtime plus profiles used by both add-ins. The bundled runtime is installed privately under the LaTeX Blocks
application directory;
the installer records that location under the current user's LaTeX Blocks registry key. `STEMTEX_HOME`, when explicitly
set, still takes precedence for development and diagnostics. The package targets 64-bit Windows and 64-bit Microsoft
Office; it also carries the matching VC++ 2015–2022 x64 runtime and installs it only when the machine's copy is older.
