param([Alias('Host')][ValidateSet('Word','PowerPoint','Both')][string]$TargetHost = 'Both')
$targets = @()
if ($TargetHost -in @('Word', 'Both')) {
    $targets += [pscustomobject]@{ Host = 'Word'; RegistryHost = 'Word'; Project = 'LaTeXBlocks.Word.AddIn' }
}
if ($TargetHost -in @('PowerPoint', 'Both')) {
    $targets += [pscustomobject]@{ Host = 'PowerPoint'; RegistryHost = 'PowerPoint'; Project = 'LaTeXBlocks.PowerPoint.AddIn' }
}
foreach ($target in $targets) {
    $key = "HKCU:\Software\Microsoft\Office\$($target.RegistryHost)\Addins\$($target.Project)"
    if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse -Force }
    Write-Output "LaTeX Blocks unregistered from $($target.Host)."
}
