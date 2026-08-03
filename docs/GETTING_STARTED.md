# Getting Started

LaTeX Blocks is a Windows desktop VSTO project. A development build is deliberately separate from the installed
product: building does not change the Office add-in registration. Register a build explicitly only when you want
Word or PowerPoint to load it.

## Prerequisites

- 64-bit Windows and 64-bit desktop Microsoft Word and/or PowerPoint. StemTeX is an x64 native runtime, so 32-bit
  Office is not supported.
- Visual Studio 18 Community with the Office/VSTO tooling and .NET Framework 4.8 targeting support. The build scripts
  currently expect MSBuild at `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`.
- The .NET Framework 4.8 and Visual Studio Tools for Office runtime to load an add-in in Office.
- PowerShell 7 (`pwsh.exe`) for the commands below.
- A usable StemTeX installation. For development and publishing, the expected staged distribution is
  `..\xetex\stemtex\dist\stemtex-installer\StemTeX` relative to this repository. It must include the x64 runtime and
  the profiles required by the package, including `xits_cjk` and `arial_lete_simhei`.

The release installer bundles StemTeX, the VSTO manifests, and the VC++ x64 runtime. A source checkout does not.

## Initialize local signing

From the repository root, run:

```powershell
pwsh.exe -NoProfile -File .\scripts\Initialize-LaTeXBlocks.ps1
```

The script creates or reuses a current-user `CN=LaTeX Blocks Development` code-signing certificate, trusts it in the
current user's `TrustedPublisher` store, and writes the certificate thumbprint to a development props file for each
add-in. Those props files are ignored by Git and must remain local.

## Build

```powershell
pwsh.exe -NoProfile -File .\scripts\Build-LaTeXBlocks.ps1 -Configuration Debug
```

Use `-Configuration Release` for an optimized local build. The script builds `LaTeXBlocks.sln` and explicitly keeps
the standard VSTO registration targets disabled. A build alone never redirects the next Office launch to
`bin\Debug` or `bin\Release`.

## Load a development build

Close the target Office applications first, then register the build you just made:

```powershell
pwsh.exe -NoProfile -File .\scripts\Register-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Both
```

`-TargetHost` accepts `Word`, `PowerPoint`, or `Both`. The registration is per-user and points the corresponding
Office host at the local `.vsto` manifest with `|vstolocal`. Reopen the selected host after registration and use the
**LaTeX Blocks** Ribbon tab.

For example, to work only on Word:

```powershell
pwsh.exe -NoProfile -File .\scripts\Register-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Word
```

Development registration intentionally replaces the matching live per-user VSTO registration. It does not copy the
build into a product location.

## Stop using the development build

Close Word and PowerPoint, then remove the local registration:

```powershell
pwsh.exe -NoProfile -File .\scripts\Unregister-LaTeXBlocks.ps1 -TargetHost Both
```

`Unregister-LaTeXBlocks.ps1` deletes the relevant add-in registry key; it does not restore an earlier installed
manifest. To return to an installed release after development registration, run that release's installer again, then
reopen Office.

## Typical edit-build-test cycle

```powershell
pwsh.exe -NoProfile -File .\scripts\Build-LaTeXBlocks.ps1 -Configuration Debug
pwsh.exe -NoProfile -File .\scripts\Register-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Both
pwsh.exe -NoProfile -File .\scripts\Test-LaTeXBlocks.ps1 -Configuration Debug -TargetHost Both
```

The test command exercises the code directly and does not itself change the Office registration. See
[Testing](TESTING.md) for the Office instances it creates and the PowerPoint resize integration test.
