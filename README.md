# LaTeX Blocks

LaTeX Blocks is a Word add-in for inserting editable LaTeX content as portable SVG objects. StemTeX performs the rendering; Microsoft Word owns only the visual object's placement and document persistence.

This repository intentionally does not implement document search. [Comprehensive Find](https://github.com/zhaiyusci/ComprehensiveFind) can search ordinary Word text and object Alternative Text as a single result space, so LaTeX source remains discoverable without hidden proxy text.

## Current prototype

- Insert a traditional inline formula at its TeX natural width.
- Insert a fixed-width LaTeX block for display, multiline, or paragraph content.
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

The smoke test has verified SVG insertion, canonical-source persistence, metadata persistence, atomic replacement,
Word-native equation numbering and bookmarks, deletion-time renumbering, and DOCX reopen in desktop Word.

## Object contract

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

- 64-bit Microsoft Word desktop
- Visual Studio with Office/VSTO development components
- .NET Framework 4.8
- a StemTeX stage containing `stemtex-renderer.dll`, `dvisvgmdaemon.dll`, and the `unicodemath_cjk` profile

`STEMTEX_HOME`, when set, selects one runtime explicitly. Otherwise the add-in examines the installed and adjacent Scholia/StemTeX development stages and selects the usable runtime with the highest semantic version. The current binding requires StemTeX 0.11's native per-request font-size API.

At Word startup the add-in discovers valid profile directories and queues creation of the renderer for the globally selected profile (`xits_cjk` when no preference has been saved). Word's UI thread is not blocked by the cold start. Like the StemTeX GUI, the add-in owns one renderer and one dedicated FIFO background thread. Renderer creation, rendering, profile replacement, and destruction all happen on that same thread. A profile switch disposes the old renderer and creates the new one; Word exit disposes the current renderer.

The editor uses a 300ms input debounce. Every request receives a monotonically increasing UI request ID. The dedicated worker skips queued requests that are no longer latest, and the UI discards any completed result whose generation or request ID has become stale. Syntax errors from automatic preview stay in the status line rather than opening modal dialogs; the Preview button remains available for an immediate refresh with an explicit error dialog.

After the latest request passes the ID check, its SVG bytes are embedded directly into a fresh preview document. The editor does not navigate to SVG files, so browser navigation state and file caching cannot leave an older formula on screen. The smoke test runs a real hidden editor message loop and verifies that changing `x_1` to `x_2+y_2` changes the committed preview SVG.

Insert and Update are enabled only when the preview exactly matches the current source, mode, width, and global profile. They embed that already-rendered SVG directly; they never repeat native rendering on Word's UI thread.

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Initialize-LaTeXBlocks.ps1
powershell -ExecutionPolicy Bypass -File scripts/Build-LaTeXBlocks.ps1 -Configuration Debug
powershell -ExecutionPolicy Bypass -File scripts/Register-LaTeXBlocks.ps1 -Configuration Debug
```

Close all Word processes before reopening Word after registration. The Ribbon tab is named **LaTeX Blocks**.

Run the end-to-end smoke test with:

```powershell
tests/LaTeXBlocks.WordSmoke/bin/Debug/LaTeXBlocks.WordSmoke.exe
```

The test executable is explicitly x64 because the current StemTeX SDK and Word installation are x64.
