# Releasing

`Publish-LaTeXBlocks.ps1` produces one self-contained per-user installer containing both VSTO add-ins and a private
StemTeX runtime. It publishes the Word and PowerPoint ClickOnce manifests, stages the package, invokes Inno Setup,
optionally Authenticode-signs the installer, and writes a SHA-256 checksum.

## Release prerequisites

- Complete the development-signing setup with `Initialize-LaTeXBlocks.ps1` under the user who will publish. The
  publishing script requires a valid current-user `CN=LaTeX Blocks Development` certificate.
- Install Visual Studio 18 Community at the location expected by the script, including MSBuild and the VC++ x64
  redistributable under its Visual Studio installation.
- Install Inno Setup 6 at `C:\Program Files (x86)\Inno Setup 6\ISCC.exe`.
- Provide a staged StemTeX distribution. By default the script uses
  `..\xetex\stemtex\dist\stemtex-installer\StemTeX`; override it with `-StemTeXSourceDir` when necessary. It must
  contain StemTeX `0.12.0` or newer, the required native renderer/daemon files, and the `xits_cjk` and
  `arial_lete_simhei` profiles.
- Close Word and PowerPoint before installing or manually checking the result.

The script clears `dist\staging`, `dist\release`, and both `bin\Release\app.publish` directories before publishing.
Do not store hand-maintained files there.

## Keep the version consistent

Choose a three-part release version such as `0.2.14`. Keep every source-of-truth entry aligned before publishing:

| Location | Required form for `0.2.14` |
| --- | --- |
| `src\LaTeXBlocks.Word.AddIn\LaTeXBlocks.Word.AddIn.csproj` | `ApplicationVersion` = `0.2.14.0` |
| `src\LaTeXBlocks.PowerPoint.AddIn\LaTeXBlocks.PowerPoint.AddIn.csproj` | `ApplicationVersion` = `0.2.14.0` |
| Both `Properties\AssemblyInfo.cs` files | `AssemblyVersion` and `AssemblyFileVersion` = `0.2.14.0` |
| `scripts\Publish-LaTeXBlocks.ps1` | default `Version` = `0.2.14` |
| `installer\LaTeXBlocks.iss` | fallback `MyAppVersion` = `0.2.14` |

The explicit `-Version` argument passed to the publish script determines the ClickOnce application version
(`0.2.14.0`) and the installer version (`0.2.14`). Updating the source entries as well keeps local About data,
development builds, and a no-argument publish consistent with the shipped package.

## Build and test the candidate

```powershell
pwsh.exe -NoProfile -File .\scripts\Build-LaTeXBlocks.ps1 -Configuration Release
pwsh.exe -NoProfile -File .\scripts\Test-LaTeXBlocks.ps1 -Configuration Release -TargetHost Both
```

The PowerPoint width integration test is a package-validation gate: run it **after** installing the generated
candidate, as described below. A VSTO solution identity has one deployment codebase, so a registry-only switch from
an installed package to `bin\Release` is not a valid release test. See [Testing](TESTING.md) for the clean-profile
development-build case.

## Publish

With the default staged StemTeX location:

```powershell
pwsh.exe -NoProfile -File .\scripts\Publish-LaTeXBlocks.ps1 -Version 0.2.14
```

With an explicit StemTeX stage:

```powershell
pwsh.exe -NoProfile -File .\scripts\Publish-LaTeXBlocks.ps1 `
  -Version 0.2.14 `
  -StemTeXSourceDir 'C:\path\to\StemTeX'
```

The expected outputs are:

```text
dist\release\LaTeXBlocks-Setup-0.2.14.exe
dist\release\LaTeXBlocks-Setup-0.2.14.exe.sha256
```

The installer bundles the Word and PowerPoint ClickOnce publications, the StemTeX runtime and profiles, its public
publisher certificate, and the VC++ 2015-2022 x64 redistributable. It performs per-user VSTO registration for both
hosts and stores the bundled runtime location in the current-user LaTeX Blocks settings.

If the Windows SDK x64 `signtool.exe` is found, the script Authenticode-signs the installer with the development
certificate. Its absence does not make the script fail, so signature verification must be an explicit release gate.

## Verify the deliverables

```powershell
$installer = '.\dist\release\LaTeXBlocks-Setup-0.2.14.exe'
Get-AuthenticodeSignature $installer | Format-List Status, StatusMessage, SignerCertificate
Get-FileHash $installer -Algorithm SHA256
Get-Content "$installer.sha256"
```

Confirm that the computed SHA-256 matches the checksum file and, when signing was expected, that the signature is
present and valid for the publishing user. Finally, install the package on a clean user registration, reopen Word and
PowerPoint, and confirm that each loads **LaTeX Blocks** and can render with the bundled StemTeX runtime. Then run:

```powershell
pwsh.exe -NoProfile -File .\tests\Test-PowerPointWidthIntegration.ps1 -Configuration Release
```

This verifies the installed PowerPoint VSTO deployment, not a registry-only development substitution.
