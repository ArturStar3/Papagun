param(
    [Parameter(Mandatory=$true)][string]$Presentation,
    [Parameter(Mandatory=$true)][string]$Image,
    [int]$Slide = 1
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $Presentation).Path
$target = [System.IO.Path]::GetFullPath($Image)
[System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($target)) | Out-Null
$app = $null
$deck = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $deck = $app.Presentations.Open($source, -1, 0, 0)
    $deck.Slides.Item($Slide).Export($target, 'PNG', 1920, 1080)
    Write-Output 'Slide exported successfully.'
} finally {
    if ($null -ne $deck) { $deck.Close(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($deck) }
    if ($null -ne $app) { $app.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
