param(
    [ValidateSet('Debug','Release')][string]$Configuration = 'Debug',
    [Alias('Host')][ValidateSet('Word','PowerPoint','Both')][string]$TargetHost = 'Both'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$targets = @()
if ($TargetHost -in @('Word', 'Both')) {
    $targets += [pscustomobject]@{
        Host = 'Word'
        Project = 'LaTeXBlocks.Word.AddIn'
        RegistryHost = 'Word'
        Process = 'WINWORD'
        Description = 'Editable, searchable LaTeX blocks rendered by StemTeX'
    }
}
if ($TargetHost -in @('PowerPoint', 'Both')) {
    $targets += [pscustomobject]@{
        Host = 'PowerPoint'
        Project = 'LaTeXBlocks.PowerPoint.AddIn'
        RegistryHost = 'PowerPoint'
        Process = 'POWERPNT'
        Description = 'Editable, searchable LaTeX blocks for slides rendered by StemTeX'
    }
}

foreach ($target in $targets) {
    $manifest = Join-Path $root "src\$($target.Project)\bin\$Configuration\$($target.Project).vsto"
    if (-not (Test-Path -LiteralPath $manifest)) {
        throw "Build LaTeX Blocks before registering $($target.Host): $manifest"
    }
    $key = "HKCU:\Software\Microsoft\Office\$($target.RegistryHost)\Addins\$($target.Project)"
    if (-not (Test-Path -LiteralPath $key)) { [void](New-Item -Path $key -Force) }
    $uri = ([Uri](Resolve-Path -LiteralPath $manifest).Path).AbsoluteUri + '|vstolocal'
    Set-ItemProperty -LiteralPath $key -Name FriendlyName -Value 'LaTeX Blocks'
    Set-ItemProperty -LiteralPath $key -Name Description -Value $target.Description
    Set-ItemProperty -LiteralPath $key -Name LoadBehavior -Type DWord -Value 3
    Set-ItemProperty -LiteralPath $key -Name Manifest -Value $uri
    Write-Output "Registered LaTeX Blocks for $($target.Host): $uri"
    if (Get-Process $target.Process -ErrorAction SilentlyContinue) {
        Write-Warning "Close every $($target.Host) window and reopen it to load this build."
    }
}
