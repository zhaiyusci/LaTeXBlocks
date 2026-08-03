# Testing

LaTeX Blocks tests are desktop smoke tests, not mocked Office tests. They require the selected Microsoft Office host
and a usable StemTeX runtime. Build the requested configuration before running them.

Close Word and PowerPoint before a test run. In particular, the PowerPoint smoke test fails immediately if any
`POWERPNT` process is already running.

## Build the test executables

```powershell
pwsh.exe -NoProfile -File .\scripts\Build-LaTeXBlocks.ps1 -Configuration Debug
```

The solution builds the Word and PowerPoint add-ins together with their smoke executables. Building and testing do
not register a VSTO add-in or replace the manifest used by an installed product.

## Standard smoke tests

Run both hosts:

```powershell
pwsh.exe -NoProfile -File .\scripts\Test-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Both
```

Run one host when iterating on host-specific code:

```powershell
pwsh.exe -NoProfile -File .\scripts\Test-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Word
pwsh.exe -NoProfile -File .\scripts\Test-LaTeXBlocks.ps1 -Configuration Debug -TargetHost PowerPoint
```

`-TargetHost` accepts `Word`, `PowerPoint`, and `Both`; `Both` runs the Word executable first and then the
PowerPoint executable.

### What each smoke test launches

| Test | Office behavior | Main coverage |
| --- | --- | --- |
| Word smoke | Starts a separate hidden Word instance and closes it when complete. It rejects an accidental attachment to an existing visible Word instance. | StemTeX startup/shutdown, rendering, SVG insertion and replacement, metadata, baseline/inline behavior, numbered equations, and DOCX persistence. |
| PowerPoint smoke | Requires PowerPoint to be closed, then starts a visible temporary PowerPoint instance and closes it when complete. | Rendering, block insertion and editing, SVG-shell style rendering/persistence, profile/TeX-size/width controls, unified host-frame resize handling, and PPTX persistence. |

The tests leave diagnostic documents under ignored `artifacts` directories. They are useful when diagnosing a failed
Office assertion but are not release artifacts.

## PowerPoint host-frame integration test

The separate PowerPoint integration test exercises the unified native host-frame contract through direct horizontal,
vertical, and corner resizes. Every size change must submit a real StemTeX layout pass: width changes update the stored
typesetting width, while height-only changes rerender the current width. It verifies that SVG content remains 1:1 with
neither stretching nor cropping, then explicitly checks that move and rotation preserve the same Shape ID and do not
rerender. It also sends two native gestures in succession after the first debounce window, so a late render from the
first may not overwrite the second. It uses the presentation created by the PowerPoint smoke test, so run that smoke
test first.

With an installed PowerPoint add-in registered:

```powershell
pwsh.exe -NoProfile -File .\tests\Test-PowerPointWidthIntegration.ps1 -Configuration Debug
```

To test a local build without leaving it registered, use a clean Windows user profile (with no installed LaTeX
Blocks PowerPoint VSTO deployment):

```powershell
pwsh.exe -NoProfile -File .\tests\Test-PowerPointWidthIntegration.ps1 -Configuration Debug -RegisterDevelopmentBuild
```

The latter mode temporarily registers the selected development build, launches PowerPoint, connects the add-in,
resizes the generated block, and restores the previous PowerPoint registration in its cleanup path. It accepts
`-ResizeFactor` (default `1.18`) and `-TimeoutSeconds` (default `30`) when investigating event timing.

Do not use this registry-only mode to replace an installed package on the same user profile. VSTO permits a solution
identity to have only one deployment codebase and will reject the switch with
`AddInAlreadyInstalledException`. For a release candidate, install the package and run the command without
`-RegisterDevelopmentBuild`; that is the authoritative PowerPoint integration test.

## Release configuration

Use the same commands with `-Configuration Release` before packaging:

```powershell
pwsh.exe -NoProfile -File .\scripts\Build-LaTeXBlocks.ps1 -Configuration Release
pwsh.exe -NoProfile -File .\scripts\Test-LaTeXBlocks.ps1 -Configuration Release -TargetHost Both
```

For a package candidate, install the intended package and run the PowerPoint width integration test against that
installed candidate.
