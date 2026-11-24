# Export slides of a deck to PNG through PowerPoint COM (faster than officecli screenshot for many slides).
#   pwsh -File scripts/ppt_export.ps1 "<deck.pptx>" "<out_dir>" [-Slides "9,12,13"] [-Width 1600]
# Works on an ASCII-named temp copy (PowerPoint can fail on long non-ASCII paths); the deck itself is not modified. Files: <out_dir>/sNN.png
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Out,
      [string]$Slides = "", [int]$Width = 1600)
$SlideList = @($Slides -split "," | Where-Object { $_ -match "\d" } | ForEach-Object { [int]$_ })
$full = (Resolve-Path $Path).Path
New-Item -ItemType Directory -Force $Out | Out-Null
$outFull = (Resolve-Path $Out).Path
$work = Join-Path $env:TEMP ("ppt_export_" + [guid]::NewGuid().ToString("N") + ".pptx")
Copy-Item -LiteralPath $full -Destination $work -Force
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($work, $true, $false, $false)
try {
    $h = [int]($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
    foreach ($slide in $pres.Slides) {
        if ($SlideList.Count -gt 0 -and $SlideList -notcontains $slide.SlideIndex) { continue }
        $f = Join-Path $outFull ("s{0:D2}.png" -f $slide.SlideIndex)
        $slide.Export($f, "PNG", $Width, $h)
    }
} finally {
    $pres.Close(); $app.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    Remove-Item -LiteralPath $work -Force -ErrorAction SilentlyContinue
}
Write-Output "exported to $outFull"
