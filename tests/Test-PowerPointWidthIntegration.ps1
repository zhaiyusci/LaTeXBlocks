param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [double]$ResizeFactor = 0.82,

    [int]$TimeoutSeconds = 30,

    [switch]$RegisterDevelopmentBuild
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$registryKey = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\LaTeXBlocks.PowerPoint.AddIn'
$presentationPath = Join-Path $PSScriptRoot 'artifacts\latex-blocks-powerpoint-smoke.pptx'
$registerScript = Join-Path $root 'scripts\Register-LaTeXBlocks.ps1'

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Close PowerPoint before running the width integration test.'
}
if (-not (Test-Path -LiteralPath $presentationPath)) {
    throw "Run the PowerPoint smoke test first: $presentationPath is missing."
}
if ($ResizeFactor -le 0 -or $ResizeFactor -ge 1) {
    throw 'ResizeFactor must be greater than zero and less than one.'
}

$hadRegistryKey = Test-Path -LiteralPath $registryKey
$savedRegistration = $null
if ($hadRegistryKey) {
    $savedRegistration = Get-ItemProperty -LiteralPath $registryKey
}

$application = $null
$presentation = $null
$selectedShape = $null
$replacementShape = $null
$addIn = $null
$testFailure = $null

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value) } catch { }
    }
}

function Get-ShapeTag([object]$Shape, [string]$Name) {
    for ($index = 1; $index -le $Shape.Tags.Count; $index++) {
        if ($Shape.Tags.Name($index) -eq $Name) {
            return [string]$Shape.Tags.Value($index)
        }
    }
    return ''
}

function ConvertFrom-LaTeXBlockTitle([string]$Title) {
    $values = @{}
    foreach ($part in $Title -split ';') {
        $separator = $part.IndexOf('=')
        if ($separator -gt 0) {
            $values[$part.Substring(0, $separator)] = $part.Substring($separator + 1)
        }
    }
    if (-not $values.ContainsKey('id') -or -not $values.ContainsKey('width')) {
        throw "Invalid LaTeX block metadata: $Title"
    }
    return [pscustomobject]@{
        id = [string]$values.id
        widthPt = [double]::Parse([string]$values.width,
            [Globalization.CultureInfo]::InvariantCulture)
    }
}

function Get-LaTeXBlockShape([object]$Presentation, [string]$StableId = '') {
    for ($slideIndex = 1; $slideIndex -le $Presentation.Slides.Count; $slideIndex++) {
        $slide = $Presentation.Slides.Item($slideIndex)
        try {
            for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
                $candidate = $slide.Shapes.Item($shapeIndex)
                $keep = $false
                try {
                    if ((Get-ShapeTag $candidate 'LATEXBLOCKS_KIND') -ne 'LATEX_BLOCK') { continue }
                    if ([string]::IsNullOrEmpty($StableId)) {
                        $keep = $true
                        return $candidate
                    }
                    $metadata = ConvertFrom-LaTeXBlockTitle $candidate.Title
                    if ([string]$metadata.id -eq $StableId) {
                        $keep = $true
                        return $candidate
                    }
                }
                finally {
                    if (-not $keep) { Release-ComObject $candidate }
                }
            }
        }
        finally {
            Release-ComObject $slide
        }
    }
    return $null
}

try {
    if ($RegisterDevelopmentBuild) {
        & $registerScript -Configuration $Configuration -TargetHost PowerPoint | Write-Host
    }
    elseif (-not $hadRegistryKey) {
        throw 'The installed PowerPoint add-in is not registered.'
    }
    Add-Type -AssemblyName System.Windows.Forms

    $application = New-Object -ComObject PowerPoint.Application
    $application.Visible = -1
    Start-Sleep -Milliseconds 1200

    for ($index = 1; $index -le $application.COMAddIns.Count; $index++) {
        $candidate = $application.COMAddIns.Item($index)
        if ($candidate.ProgId -eq 'LaTeXBlocks.PowerPoint.AddIn') {
            $addIn = $candidate
            break
        }
        Release-ComObject $candidate
    }
    if ($null -eq $addIn) {
        throw 'PowerPoint did not expose the LaTeX Blocks COM add-in.'
    }
    if (-not $addIn.Connect) {
        $addIn.Connect = $true
        Start-Sleep -Milliseconds 1200
    }
    if (-not $addIn.Connect) {
        throw 'PowerPoint did not connect the LaTeX Blocks add-in.'
    }

    $presentation = $application.Presentations.Open($presentationPath, 0, 0, 0)
    $selectedShape = Get-LaTeXBlockShape $presentation
    if ($null -eq $selectedShape) {
        throw 'The smoke presentation contains no recognized LaTeX block.'
    }

    $before = ConvertFrom-LaTeXBlockTitle $selectedShape.Title
    $stableId = [string]$before.id
    $layoutWidthBefore = [double]$before.widthPt
    $visibleWidthBefore = [double]$selectedShape.Width
    $selectedShape.Width = [single]($visibleWidthBefore * $ResizeFactor)

    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $after = $null
    while ([DateTime]::UtcNow -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
        $candidate = Get-LaTeXBlockShape $presentation $stableId
        if ($null -eq $candidate) { continue }
        try {
            $metadata = ConvertFrom-LaTeXBlockTitle $candidate.Title
            if ([Math]::Abs([double]$metadata.widthPt - $layoutWidthBefore * $ResizeFactor) -lt 0.15) {
                $replacementShape = $candidate
                $candidate = $null
                $after = $metadata
                break
            }
        }
        finally {
            Release-ComObject $candidate
        }
    }
    if ($null -eq $replacementShape) {
        throw "The resize event did not commit a width update within $TimeoutSeconds seconds."
    }

    $svgWidth = [double](Get-ShapeTag $replacementShape 'LATEXBLOCKS_SVG_WIDTH_PT')
    $svgHeight = [double](Get-ShapeTag $replacementShape 'LATEXBLOCKS_SVG_HEIGHT_PT')
    $visualScale = [double](Get-ShapeTag $replacementShape 'LATEXBLOCKS_VISUAL_SCALE')
    $expectedVisibleWidth = $svgWidth * $visualScale
    $expectedVisibleHeight = $svgHeight * $visualScale
    $visibleWidthAfter = [double]$replacementShape.Width
    $visibleHeightAfter = [double]$replacementShape.Height

    if ([Math]::Abs($visibleWidthAfter - $expectedVisibleWidth) -gt 0.12 -or
        [Math]::Abs($visibleHeightAfter - $expectedVisibleHeight) -gt 0.12) {
        throw "The replacement is distorted: actual=$visibleWidthAfter x $visibleHeightAfter; expected=$expectedVisibleWidth x $expectedVisibleHeight."
    }

    [pscustomobject]@{
        AddInConnected = [bool]$addIn.Connect
        LayoutWidthBeforePt = [Math]::Round($layoutWidthBefore, 3)
        LayoutWidthAfterPt = [Math]::Round([double]$after.widthPt, 3)
        RequestedFactor = $ResizeFactor
        VisibleWidthAfterPt = [Math]::Round($visibleWidthAfter, 3)
        ExpectedVisibleWidthPt = [Math]::Round($expectedVisibleWidth, 3)
        VisibleHeightAfterPt = [Math]::Round($visibleHeightAfter, 3)
        ExpectedVisibleHeightPt = [Math]::Round($expectedVisibleHeight, 3)
    }
}
catch {
    $testFailure = $_
}
finally {
    Release-ComObject $replacementShape
    Release-ComObject $selectedShape
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { }
        Release-ComObject $presentation
    }
    Release-ComObject $addIn
    if ($null -ne $application) {
        try { $application.Quit() } catch { }
        Release-ComObject $application
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($hadRegistryKey) {
        if (-not (Test-Path -LiteralPath $registryKey)) {
            [void](New-Item -Path $registryKey -Force)
        }
        foreach ($name in 'FriendlyName', 'Description', 'Manifest', 'LoadBehavior') {
            if ($null -ne $savedRegistration.$name) {
                Set-ItemProperty -LiteralPath $registryKey -Name $name -Value $savedRegistration.$name
            }
        }
    }
    elseif (Test-Path -LiteralPath $registryKey) {
        Remove-Item -LiteralPath $registryKey -Recurse -Force
    }
}

if ($null -ne $testFailure) { throw $testFailure }
