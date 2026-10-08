# Export the saved draft through Microsoft PowerPoint for an actual render check.
# Open read-only without a document window; close only the presentation we opened.
param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$PdfPath,
    [Parameter(Mandatory=$true)][string]$PreviewDir
)
$ErrorActionPreference = 'Stop'
$powerpoint = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $powerpoint.Presentations.Open($SourcePath, -1, 0, 0)
    # Include production appendices in the review PDF only; source stays read-only.
    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $presentation.Slides.Item($slideIndex).SlideShowTransition.Hidden = 0
    }
    $presentation.SaveAs($PdfPath, 32)
    New-Item -ItemType Directory -Force -Path $PreviewDir | Out-Null
    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $imagePath = Join-Path $PreviewDir ('slide-{0:D2}.png' -f $slideIndex)
        $presentation.Slides.Item($slideIndex).Export($imagePath, 'PNG', 1600, 900)
    }
    Write-Output ('Rendered {0} slides with Microsoft PowerPoint.' -f $presentation.Slides.Count)
}
finally {
    if ($null -ne $presentation) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    # Do not quit the application; an existing user-owned session may be active.
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint)
}
