# StemTeX Integration

LaTeX Blocks is an Office integration layer, not a TeX distribution. [StemTeX](https://github.com/zhaiyusci/StemTeX)
is a separate repository with its own build and release lifecycle. It is **not** a Git submodule here: this repository
consumes a staged StemTeX runtime at development and release time. The x64 native renderer is loaded only by the
external `LaTeXBlocks.RenderHost.host` process, never by Word or PowerPoint's VSTO AppDomain. Existing Office
documents remain displayable without it because they embed the already-rendered SVG; a runtime is required only to
preview, insert, edit, or rerender a block.

The host split and rendering ownership are described in [Architecture](ARCHITECTURE.md). Word's SVG and TeX-metric
contract is specified in the [Word object model](OBJECT_MODEL.md); PowerPoint deliberately consumes the same renderer
only for independent slide blocks, as described in [PowerPoint scope](POWERPOINT_SCOPE.md).

## Supported runtime contract

LaTeX Blocks requires StemTeX **0.12.0 or newer**. A usable runtime root has all of the following:

```text
<StemTeXHome>/
  runtime/VERSION                         # parseable semantic version, >= 0.12.0
  runtime/bin/sdk/stemtex-renderer.dll
  runtime/bin/windows/dvisvgmdaemon.dll
  gui/profiles/<profile>/preamble.tex
```

The version gate is necessary but not sufficient. `stemtex-renderer.dll` must be an x64 build compatible with the
0.12 native ABI and export these Cdecl entry points. They are loaded by `LaTeXBlocks.RenderHost.host`, not by an
Office add-in:

- `stemtex_renderer_create`
- `stemtex_renderer_render_output_bytes_with_font_size`
- `stemtex_renderer_free_output_bytes`
- `stemtex_renderer_free_output_result`
- `stemtex_renderer_free_string`
- `stemtex_renderer_cancel_current`
- `stemtex_renderer_destroy`

The render host supplies the runtime root and selected profile directory when creating the renderer. Each pipe render
request supplies the TeX source, typesetting width, auto-width flag, and real TeX design size in points. The
per-request font-size API is important: the Office host never fakes a new TeX size by scaling a previously rendered
SVG. The RenderHost returns self-contained SVG bytes plus outcome metadata; the add-in consumes the returned TeX box
metrics to place the SVG in the Office document.

## Runtime discovery

Each Office host resolves one runtime configuration when its RenderHost client is created, then launches a private
RenderHost process that uses that same current-user environment and registry configuration. `STEMTEX_HOME` is an
explicit override: set it to the runtime root **before starting Word or PowerPoint**. If it is set but does not
satisfy the contract above, startup fails; the add-in intentionally does not fall back to another installation.

Without `STEMTEX_HOME`, the add-in examines these candidates in order:

1. `HKCU\Software\LaTeXBlocks\StemTeXHome` (written by the LaTeX Blocks installer);
2. `%SystemRoot%\StemTeX`;
3. `%ProgramFiles%\Scholia\StemTeX`;
4. `%ProgramFiles%\StemTeX`;
5. `%USERPROFILE%\Documents\xetex\stemtex\dist\stemtex-installer\StemTeX`;
6. `%USERPROFILE%\Documents\xetex\stemtex\build\stemtex-check-stage`.

Candidates missing the requested profile or any required file are ignored. Among usable candidates, the highest
`runtime/VERSION` wins; ties retain the order above, so the private installed runtime wins over an equal-version
development stage. A newer development stage can therefore win deliberately, while `STEMTEX_HOME` remains the
deterministic choice for debugging a specific build.

## Profiles and host preferences

Profiles are discovered from immediate children of `gui/profiles/` that contain `preamble.tex`. The normal default is
`xits_cjk`; if that profile is absent, the first discovered profile in case-insensitive alphabetical order is used.

A profile is a **host-session preference**, not document or object metadata:

| Host | Preference location |
| --- | --- |
| Word | `HKCU\Software\LaTeXBlocks\Word\Profile` |
| PowerPoint | `HKCU\Software\LaTeXBlocks\PowerPoint\Profile` |

For migration only, a missing host-specific value may read the former shared
`HKCU\Software\LaTeXBlocks\Profile` value. An unavailable saved profile falls back to the default. Choosing a profile
updates only that host's preference and affects later preview, insertion, and rerender operations in that host. It
does not rewrite existing SVGs or change the other Office host. Blocks store their source and layout facts, not a
profile name; an explicit edit or refresh uses the then-current host profile.

## Development stages and released packages

For local development, build StemTeX separately and point `STEMTEX_HOME` at its staged distribution when a specific
runtime is required. The release script's default stage is:

```text
..\xetex\stemtex\dist\stemtex-installer\StemTeX
```

relative to this repository. It can be overridden with
`scripts\Publish-LaTeXBlocks.ps1 -StemTeXSourceDir <stage>`. Publishing validates the native files, version, and the
`xits_cjk` and `arial_lete_simhei` profiles before producing an installer.

The installer copies the selected stage's `runtime/` and `gui/profiles/` into its private
`%LocalAppData%\Programs\LaTeX Blocks\StemTeX` directory and writes that location to
`HKCU\Software\LaTeXBlocks\StemTeXHome`. Thus an installed product is self-contained with a fixed renderer/profile
set; it does not require a sibling StemTeX source checkout. Runtime cache and intermediate artifacts are excluded from
the package. The x64 `LaTeXBlocks.RenderHost.host` is delivered beside the add-ins so each Office host can launch its
private native-rendering process. See the publishing script and installer definition for the exact release inputs:
[Publish-LaTeXBlocks.ps1](../scripts/Publish-LaTeXBlocks.ps1) and
[LaTeXBlocks.iss](../installer/LaTeXBlocks.iss).

## RenderHost process and asynchronous request boundary

Each Office host owns a `RenderHostClientBackend`, selected profile preference, and one private
`LaTeXBlocks.RenderHost.host` child process. The add-in starts it with a unique pipe nonce and sends
length-framed JSON requests over a current-user-only named pipe. The native `StemTeXBackend` and
`stemtex-renderer.dll` exist only inside the x64 RenderHost; Word and PowerPoint retain no in-process native renderer
or shutdown reaper. Office COM mutation remains on the corresponding host UI path.

- `switchProfile` queues warm-up and returns immediately. Add-in startup and profile changes therefore do not block
  the Office UI; a subsequent render waits behind the same RenderHost queue if warm-up is still in progress.
- Live editor previews use `renderLatest`. A newer preview, or `cancelLatestPreview` sent on another pipe connection,
  supersedes the old preview; the client maps that response to normal task cancellation rather than a user-visible
  renderer error.
- Insertions, edits, width changes, font-size refreshes, and native host-frame reflows use `renderQueued`. These are
  durable FIFO work: a later preview cannot silently cancel a user command.
- Profile changes advance the RenderHost backend generation so older work becomes stale before a replacement renderer
  is accepted.
- The Office-side client assigns its RenderHost process to a Windows Job Object with `KILL_ON_JOB_CLOSE`. On add-in
  unload it closes that job handle and returns immediately; the render host and its descendants are terminated without
  waiting for native creation, rendering, cancellation, or disposal.

This boundary mirrors StemTeX GUI's responsiveness model while preserving the Office rule that all document changes
belong to the host thread. It is a renderer lifecycle policy, not a document watcher; see [Architecture](ARCHITECTURE.md)
and the rationale in [Design decisions](DECISIONS.md).

## Diagnosing a missing runtime

If the add-in reports that StemTeX with SVG support was not found, first check `STEMTEX_HOME`. An explicitly set but
invalid value blocks all fallback discovery. Otherwise inspect the installed registry value and confirm that the chosen
root has a compatible `runtime/VERSION`, both native DLLs, and a `preamble.tex` for the selected profile. The Word
runtime diagnostics surface the resolved StemTeX home, available profiles, current profile, and backend status.
