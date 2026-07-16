param([ValidateSet('Debug','Release')][string]$Configuration='Debug')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$manifest=Join-Path $root "src\LaTeXBlocks.Word.AddIn\bin\$Configuration\LaTeXBlocks.Word.AddIn.vsto"
if(-not (Test-Path $manifest)){throw 'Build LaTeX Blocks before registering it.'}
$key='HKCU:\Software\Microsoft\Office\Word\Addins\LaTeXBlocks.Word.AddIn'
if(-not (Test-Path $key)){[void](New-Item $key -Force)}
$uri=([Uri](Resolve-Path $manifest).Path).AbsoluteUri+'|vstolocal'
Set-ItemProperty $key FriendlyName 'LaTeX Blocks'
Set-ItemProperty $key Description 'Editable, searchable LaTeX blocks rendered by StemTeX'
Set-ItemProperty $key LoadBehavior -Type DWord 3
Set-ItemProperty $key Manifest $uri
Write-Output "Registered LaTeX Blocks: $uri"
if(Get-Process WINWORD -ErrorAction SilentlyContinue){Write-Warning 'Close every Word window and reopen Word to load this build.'}
