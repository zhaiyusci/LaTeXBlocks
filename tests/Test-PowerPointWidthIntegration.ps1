param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [double]$ResizeFactor = 1.18,

    [double]$CornerResizeFactor = 1.10,

    [double]$AutoReflowFactor = 0.65,

    [double]$FrameHeightPaddingPt = 36,

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
if ($ResizeFactor -le 0 -or $ResizeFactor -gt 5) {
    throw 'ResizeFactor must be greater than zero and no greater than five.'
}
if ($CornerResizeFactor -le 0 -or $CornerResizeFactor -gt 5) {
    throw 'CornerResizeFactor must be greater than zero and no greater than five.'
}
if ($AutoReflowFactor -le 0 -or $AutoReflowFactor -ge 1) {
    throw 'AutoReflowFactor must be greater than zero and less than one.'
}
if ($FrameHeightPaddingPt -le 0 -or $FrameHeightPaddingPt -gt 360) {
    throw 'FrameHeightPaddingPt must be greater than zero and no greater than 360 pt.'
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
$shrinkShape = $null
$reflowShape = $null
$verticalShape = $null
$cornerShape = $null
$nonReflowShape = $null
$rapidResizeShape = $null
$lastGestureShape = $null
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
    if (-not $values.ContainsKey('id') -or -not $values.ContainsKey('width') -or
        -not $values.ContainsKey('size')) {
        throw "Invalid LaTeX block metadata: $Title"
    }
    return [pscustomobject]@{
        id = [string]$values.id
        widthPt = [double]::Parse([string]$values.width,
            [Globalization.CultureInfo]::InvariantCulture)
        fontSizePt = [double]::Parse([string]$values.size,
            [Globalization.CultureInfo]::InvariantCulture)
    }
}

function ConvertTo-InvariantDouble([string]$Text, [string]$Description) {
    if ([string]::IsNullOrWhiteSpace($Text)) {
        throw "$Description is missing."
    }
    try {
        return [double]::Parse($Text, [Globalization.NumberStyles]::Float,
            [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw "$Description is not an invariant point value: $Text"
    }
}

function Assert-Near([double]$Actual, [double]$Expected, [double]$Tolerance,
    [string]$Description) {
    if ([Math]::Abs($Actual - $Expected) -gt $Tolerance) {
        throw "${Description}: actual=$Actual, expected=$Expected, tolerance=$Tolerance."
    }
}

function Assert-HostFrameContract([object]$Shape, [object]$ExpectedMetadata,
    [double]$ExpectedFrameWidth, [double]$ExpectedFrameHeight, [string]$Label) {
    $visualScale = Get-ShapeTag $Shape 'LATEXBLOCKS_VISUAL_SCALE'
    if (-not [string]::IsNullOrEmpty($visualScale)) {
        throw "$Label retained obsolete LATEXBLOCKS_VISUAL_SCALE=$visualScale."
    }

    $svgWidthText = Get-ShapeTag $Shape 'LATEXBLOCKS_SVG_WIDTH_PT'
    $svgHeightText = Get-ShapeTag $Shape 'LATEXBLOCKS_SVG_HEIGHT_PT'
    $svgWidth = ConvertTo-InvariantDouble $svgWidthText "$Label SVG width tag"
    $svgHeight = ConvertTo-InvariantDouble $svgHeightText "$Label SVG height tag"
    $actualWidth = [double]$Shape.Width
    $actualHeight = [double]$Shape.Height

    # The tags describe the final SVG root box. They must match PowerPoint's
    # picture box: otherwise PowerPoint is still applying a host-side scale.
    Assert-Near $actualWidth $svgWidth 0.12 "$Label shape/SVG width"
    Assert-Near $actualHeight $svgHeight 0.12 "$Label shape/SVG height"
    Assert-Near $actualWidth $ExpectedFrameWidth 0.15 "$Label frame width"
    Assert-Near $actualHeight $ExpectedFrameHeight 0.15 "$Label frame height"

    $metadata = ConvertFrom-LaTeXBlockTitle $Shape.Title
    Assert-Near ([double]$metadata.widthPt) ([double]$ExpectedMetadata.widthPt) 0.15 "$Label typesetting width"
    Assert-Near ([double]$metadata.fontSizePt) ([double]$ExpectedMetadata.fontSizePt) 0.001 "$Label TeX font size"
    return $metadata
}

function Test-HostFrameGeometry([object]$Shape, [double]$ExpectedWidth,
    [double]$ExpectedHeight) {
    try {
        $svgWidth = ConvertTo-InvariantDouble (Get-ShapeTag $Shape 'LATEXBLOCKS_SVG_WIDTH_PT') 'SVG width tag'
        $svgHeight = ConvertTo-InvariantDouble (Get-ShapeTag $Shape 'LATEXBLOCKS_SVG_HEIGHT_PT') 'SVG height tag'
        return [Math]::Abs([double]$Shape.Width - $ExpectedWidth) -le 0.15 -and
               [Math]::Abs([double]$Shape.Height - $ExpectedHeight) -le 0.15 -and
               [Math]::Abs([double]$Shape.Width - $svgWidth) -le 0.12 -and
               [Math]::Abs([double]$Shape.Height - $svgHeight) -le 0.12
    }
    catch {
        return $false
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
                    if ([string]::IsNullOrWhiteSpace([string]$candidate.Title)) { continue }
                    # Shape replacement is atomic from the user's point of view,
                    # but COM enumeration can briefly expose a just-created shape
                    # before its Title property has propagated. Skip that transient
                    # candidate and keep polling for the stable semantic ID.
                    try {
                        $metadata = ConvertFrom-LaTeXBlockTitle $candidate.Title
                    }
                    catch {
                        continue
                    }
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

function Wait-ForLaTeXBlockShape([object]$Presentation, [string]$StableId,
    [int]$TimeoutMilliseconds = 1500) {
    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        $candidate = Get-LaTeXBlockShape $Presentation $StableId
        if ($null -ne $candidate) { return $candidate }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 15
    }
    return $null
}

function Wait-ForHostFrame([object]$Presentation, [string]$StableId,
    [double]$ExpectedWidthPt, [double]$ExpectedHeightPt,
    [double]$ExpectedLayoutWidthPt, [double]$ExpectedFontSizePt,
    [int]$TimeoutSeconds, [string]$Description) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
        $candidate = Get-LaTeXBlockShape $Presentation $StableId
        if ($null -eq $candidate) { continue }
        $keep = $false
        try {
            $metadata = ConvertFrom-LaTeXBlockTitle $candidate.Title
            if ([Math]::Abs([double]$metadata.widthPt - $ExpectedLayoutWidthPt) -gt 0.15 -or
                [Math]::Abs([double]$metadata.fontSizePt - $ExpectedFontSizePt) -gt 0.001 -or
                -not (Test-HostFrameGeometry $candidate $ExpectedWidthPt $ExpectedHeightPt)) {
                continue
            }
            $keep = $true
            return $candidate
        }
        finally {
            if (-not $keep) { Release-ComObject $candidate }
        }
    }
    throw "$Description did not commit its complete host frame within $TimeoutSeconds seconds."
}

function Wait-ForReflowedFrame([object]$Presentation, [string]$StableId,
    [double]$TargetFrameWidthPt, [double]$PreviousLayoutWidthPt,
    [double]$ExpectedFontSizePt,
    [ValidateSet('Increase', 'Decrease')][string]$LayoutDirection,
    [int]$TimeoutSeconds, [string]$Description) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
        $candidate = Get-LaTeXBlockShape $Presentation $StableId
        if ($null -eq $candidate) { continue }
        $keep = $false
        try {
            $metadata = ConvertFrom-LaTeXBlockTitle $candidate.Title
            $layoutChangedInExpectedDirection = if ($LayoutDirection -eq 'Increase') {
                [double]$metadata.widthPt -gt $PreviousLayoutWidthPt + 0.15
            }
            else {
                [double]$metadata.widthPt -lt $PreviousLayoutWidthPt - 0.15
            }
            if ([Math]::Abs([double]$candidate.Width - $TargetFrameWidthPt) -gt 0.15 -or
                -not $layoutChangedInExpectedDirection -or
                [Math]::Abs([double]$metadata.fontSizePt - $ExpectedFontSizePt) -gt 0.001 -or
                -not (Test-HostFrameGeometry $candidate $TargetFrameWidthPt ([double]$candidate.Height))) {
                continue
            }
            $keep = $true
            return $candidate
        }
        finally {
            if (-not $keep) { Release-ComObject $candidate }
        }
    }
    throw "$Description did not commit a reflowed TeX layout within $TimeoutSeconds seconds."
}

function Set-CornerFrame([object]$Presentation, [object]$Shape, [string]$StableId,
    [double]$WidthPt, [double]$HeightPt) {
    # PowerPoint raises AfterShapeSizeChange for each COM setter. Do not pump the
    # UI between the two writes: the add-in must coalesce them into one final
    # frame even if it sees two host notifications.
    try {
        $Shape.Width = [single]$WidthPt
        $Shape.Height = [single]$HeightPt
        return
    }
    catch {
        # A very fast replacement can invalidate the original COM proxy between
        # the two setters. Reacquire by semantic ID and finish the same frame.
        $replacement = Get-LaTeXBlockShape $Presentation $StableId
        if ($null -eq $replacement) { throw }
        try {
            $replacement.Width = [single]$WidthPt
            $replacement.Height = [single]$HeightPt
        }
        finally {
            Release-ComObject $replacement
        }
    }
}

function Set-HostFrameWidth([object]$Presentation, [object]$Shape, [string]$StableId,
    [double]$WidthPt) {
    try {
        $Shape.Width = [single]$WidthPt
        return
    }
    catch {
        # A prior native-frame render may replace the picture between gestures.
        # Reacquire by the persistent block ID, then apply the user's latest drag.
        $replacement = Get-LaTeXBlockShape $Presentation $StableId
        if ($null -eq $replacement) { throw }
        try {
            $replacement.Width = [single]$WidthPt
        }
        finally {
            Release-ComObject $replacement
        }
    }
}

function Pump-PowerPointForMilliseconds([int]$Milliseconds) {
    $deadline = [DateTime]::UtcNow.AddMilliseconds($Milliseconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 15
    }
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
    $fontSizeBefore = [double]$before.fontSizePt
    $visibleWidthBefore = [double]$selectedShape.Width
    $visibleHeightBefore = [double]$selectedShape.Height
    $horizontalTargetWidth = $visibleWidthBefore * $ResizeFactor
    $selectedShape.Width = [single]($visibleWidthBefore * $ResizeFactor)

    # Every host-frame size change is a true TeX layout request. This expansion
    # must therefore change the stored StemTeX measure rather than merely adding
    # transparent picture padding around the old SVG.
    $replacementShape = Wait-ForReflowedFrame $presentation $stableId $horizontalTargetWidth $layoutWidthBefore $fontSizeBefore 'Increase' $TimeoutSeconds 'Horizontal resize'
    $after = ConvertFrom-LaTeXBlockTitle $replacementShape.Title
    $visibleWidthAfter = [double]$replacementShape.Width
    $visibleHeightAfter = [double]$replacementShape.Height
    $null = Assert-HostFrameContract $replacementShape $after $horizontalTargetWidth $visibleHeightAfter 'Horizontal frame'
    $currentLayoutWidth = [double]$after.widthPt

    # Shrinking an already-expanded frame is another TeX layout request, not a
    # request to reclaim transparent padding from the old SVG.
    $shrunkenTargetWidth = $visibleWidthBefore + (($horizontalTargetWidth - $visibleWidthBefore) * 0.5)
    Set-HostFrameWidth $presentation $replacementShape $stableId $shrunkenTargetWidth
    Release-ComObject $replacementShape
    $replacementShape = $null

    $shrinkShape = Wait-ForReflowedFrame $presentation $stableId $shrunkenTargetWidth $currentLayoutWidth $fontSizeBefore 'Decrease' $TimeoutSeconds 'Shrink expanded frame'
    $shrinkMetadata = ConvertFrom-LaTeXBlockTitle $shrinkShape.Title
    $null = Assert-HostFrameContract $shrinkShape $shrinkMetadata $shrunkenTargetWidth ([double]$shrinkShape.Height) 'Shrunk frame'
    $visibleWidthAfter = [double]$shrinkShape.Width
    $visibleHeightAfter = [double]$shrinkShape.Height
    $currentLayoutWidth = [double]$shrinkMetadata.widthPt

    # A smaller-than-natural width is not a crop or image zoom. The add-in must
    # ask TeX for a narrower line measure, then use the resulting (possibly
    # taller) TeX box inside the requested external frame.
    $reflowTargetWidth = [Math]::Max(36.0, $visibleWidthBefore * $AutoReflowFactor)
    Set-HostFrameWidth $presentation $shrinkShape $stableId $reflowTargetWidth
    Release-ComObject $shrinkShape
    $shrinkShape = $null

    $reflowShape = Wait-ForReflowedFrame $presentation $stableId $reflowTargetWidth $currentLayoutWidth $fontSizeBefore 'Decrease' $TimeoutSeconds 'Narrow reflow'
    $reflowMetadata = ConvertFrom-LaTeXBlockTitle $reflowShape.Title
    $null = Assert-HostFrameContract $reflowShape $reflowMetadata $reflowTargetWidth ([double]$reflowShape.Height) 'Auto-reflow frame'
    $visibleWidthAfter = [double]$reflowShape.Width
    $visibleHeightAfter = [double]$reflowShape.Height
    $currentLayoutWidth = [double]$reflowMetadata.widthPt

    # A pure vertical drag must use precisely the same external-frame contract.
    $verticalTargetHeight = [Math]::Max($visibleHeightAfter + $FrameHeightPaddingPt, 120.0)
    $reflowShapeId = [int]$reflowShape.Id
    $reflowShape.Height = [single]$verticalTargetHeight
    Release-ComObject $reflowShape
    $reflowShape = $null

    $verticalShape = Wait-ForHostFrame $presentation $stableId $visibleWidthAfter $verticalTargetHeight $currentLayoutWidth $fontSizeBefore $TimeoutSeconds 'Vertical resize'
    $verticalExpected = [pscustomobject]@{
        widthPt = $currentLayoutWidth
        fontSizePt = $fontSizeBefore
    }
    $null = Assert-HostFrameContract $verticalShape $verticalExpected $visibleWidthAfter $verticalTargetHeight 'Vertical frame'
    if ([int]$verticalShape.Id -eq $reflowShapeId) {
        throw 'A vertical size change retained the old SVG instead of submitting a TeX layout pass.'
    }
    $visibleWidthAfter = [double]$verticalShape.Width
    $visibleHeightAfter = [double]$verticalShape.Height

    # A corner resize follows the same reflow contract and changes the TeX
    # measure because its target width changes too.
    $cornerTargetWidth = $visibleWidthAfter * $CornerResizeFactor
    $cornerTargetHeight = [Math]::Max($visibleHeightAfter + $FrameHeightPaddingPt, 120.0)
    Set-CornerFrame $presentation $verticalShape $stableId $cornerTargetWidth $cornerTargetHeight
    Release-ComObject $verticalShape
    $verticalShape = $null

    $cornerShape = Wait-ForReflowedFrame $presentation $stableId $cornerTargetWidth $currentLayoutWidth $fontSizeBefore 'Increase' $TimeoutSeconds 'Corner resize'
    $cornerExpected = ConvertFrom-LaTeXBlockTitle $cornerShape.Title
    $null = Assert-HostFrameContract $cornerShape $cornerExpected $cornerTargetWidth $cornerTargetHeight 'Corner frame'
    $currentLayoutWidth = [double]$cornerExpected.widthPt

    # Position and rotation are ordinary PowerPoint operations, not layout
    # operations. Neither may replace the SVG or submit a StemTeX request.
    $cornerShapeId = [int]$cornerShape.Id
    $movedLeft = [double]$cornerShape.Left + 17.0
    $rotatedDegrees = (([double]$cornerShape.Rotation + 11.0) % 360.0)
    $cornerShape.Left = [single]$movedLeft
    $cornerShape.Rotation = [single]$rotatedDegrees
    # Give delayed host events ample time to surface. Reacquire only a stable,
    # fully-contracted block; a stale COM proxy may refer to a just-deleted
    # replacement shape while PowerPoint completes its notifications.
    Pump-PowerPointForMilliseconds 2200
    $nonReflowShape = Wait-ForHostFrame $presentation $stableId $cornerTargetWidth $cornerTargetHeight $currentLayoutWidth $fontSizeBefore 5 'Move/rotation stability'
    $nonReflowMetadata = ConvertFrom-LaTeXBlockTitle $nonReflowShape.Title
    if ([int]$nonReflowShape.Id -ne $cornerShapeId) {
        throw 'Moving or rotating a LaTeX block incorrectly replaced its SVG.'
    }
    Assert-Near ([double]$nonReflowShape.Left) $movedLeft 0.15 'Moved block left position'
    Assert-Near ([double]$nonReflowShape.Rotation) $rotatedDegrees 0.15 'Rotated block angle'
    $null = Assert-HostFrameContract $nonReflowShape $nonReflowMetadata $cornerTargetWidth $cornerTargetHeight 'Move/rotation frame'
    Assert-Near ([double]$nonReflowMetadata.widthPt) $currentLayoutWidth 0.001 'Move/rotation typesetting width'
    Release-ComObject $cornerShape
    $cornerShape = $nonReflowShape
    $nonReflowShape = $null

    # The first drag is allowed to pass the add-in's short native-gesture quiet
    # period. The second drag is then issued before waiting for a completed
    # render. The final host frame must reflect the latest user intent, not a
    # stale replacement associated with the first drag.
    $firstGestureTargetWidth = $cornerTargetWidth * 1.04
    $lastGestureTargetWidth = $cornerTargetWidth * 1.08
    Set-HostFrameWidth $presentation $cornerShape $stableId $firstGestureTargetWidth
    Pump-PowerPointForMilliseconds 180
    $rapidResizeShape = Wait-ForLaTeXBlockShape $presentation $stableId
    if ($null -eq $rapidResizeShape) {
        throw 'The LaTeX block disappeared between consecutive native resize gestures.'
    }
    Assert-Near ([double]$rapidResizeShape.Width) $firstGestureTargetWidth 0.15 'First consecutive native frame width'
    Assert-Near ([double]$rapidResizeShape.Height) $cornerTargetHeight 0.15 'First consecutive native frame height'
    Set-HostFrameWidth $presentation $rapidResizeShape $stableId $lastGestureTargetWidth
    Release-ComObject $rapidResizeShape
    $rapidResizeShape = $null

    $lastGestureShape = Wait-ForReflowedFrame $presentation $stableId $lastGestureTargetWidth $currentLayoutWidth $fontSizeBefore 'Increase' $TimeoutSeconds 'Consecutive native resize'
    $lastGestureExpected = ConvertFrom-LaTeXBlockTitle $lastGestureShape.Title
    $null = Assert-HostFrameContract $lastGestureShape $lastGestureExpected $lastGestureTargetWidth $cornerTargetHeight 'Consecutive native frame'
    Release-ComObject $lastGestureShape
    $lastGestureShape = $null

    Release-ComObject $cornerShape
    $cornerShape = $null

    [pscustomobject]@{
        AddInConnected = [bool]$addIn.Connect
        LayoutWidthBeforePt = [Math]::Round($layoutWidthBefore, 3)
        LayoutWidthAfterHorizontalPt = [Math]::Round([double]$after.widthPt, 3)
        RequestedFactor = $ResizeFactor
        VisibleWidthAfterPt = [Math]::Round($visibleWidthAfter, 3)
        VisibleHeightAfterPt = [Math]::Round($visibleHeightAfter, 3)
        ShrunkFrameWidthPt = [Math]::Round($shrunkenTargetWidth, 3)
        AutoReflowFrameWidthPt = [Math]::Round($reflowTargetWidth, 3)
        AutoReflowLayoutWidthPt = [Math]::Round([double]$reflowMetadata.widthPt, 3)
        VerticalFrameHeightPt = [Math]::Round($verticalTargetHeight, 3)
        VerticalTriggeredReflow = $true
        CornerLayoutWidthPt = [Math]::Round([double]$cornerExpected.widthPt, 3)
        ConsecutiveLayoutWidthPt = [Math]::Round([double]$lastGestureExpected.widthPt, 3)
        CornerFrameWidthPt = [Math]::Round($cornerTargetWidth, 3)
        CornerFrameHeightPt = [Math]::Round($cornerTargetHeight, 3)
        MoveRotationPreserved = $true
        FirstConsecutiveFrameWidthPt = [Math]::Round($firstGestureTargetWidth, 3)
        ConsecutiveFrameWidthPt = [Math]::Round($lastGestureTargetWidth, 3)
        TeXFontSizePt = [Math]::Round($fontSizeBefore, 3)
    }
}
catch {
    $testFailure = $_
}
finally {
    Release-ComObject $shrinkShape
    Release-ComObject $reflowShape
    Release-ComObject $verticalShape
    Release-ComObject $cornerShape
    Release-ComObject $nonReflowShape
    Release-ComObject $rapidResizeShape
    Release-ComObject $lastGestureShape
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
