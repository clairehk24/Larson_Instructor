$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$workspace = (Get-Location).Path
$wordPath = 'C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE'
$inputPath = (Resolve-Path 'assets\downloads\simulation-6\instructions-for-sp-anaphylaxis.docx').Path
$outputDir = Join-Path $workspace '.codex-qa\sim6-docx-review\word-pages'
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$process = Start-Process -FilePath $wordPath -ArgumentList '/x','/automation' -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 5
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputPath, $false, $true)
    $pageCount = $document.ComputeStatistics(2)
    Write-Output "PAGES=$pageCount"

    for ($page = 1; $page -le $pageCount; $page++) {
        $start = $document.GoTo(1, 1, $page).Start
        if ($page -lt $pageCount) {
            $end = $document.GoTo(1, 1, $page + 1).Start - 1
        } else {
            $end = $document.Content.End - 1
        }
        $range = $document.Range($start, $end)
        [System.Windows.Forms.Clipboard]::Clear()
        $range.CopyAsPicture()
        Start-Sleep -Milliseconds 500
        $image = [System.Windows.Forms.Clipboard]::GetImage()
        if ($null -eq $image) {
            throw "Word did not place page $page on the clipboard as an image"
        }
        $outputPath = Join-Path $outputDir ("page-{0}.png" -f $page)
        $image.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $image.Dispose()
        Write-Output $outputPath
    }

    $document.Close(0)
    $word.Quit(0)
}
finally {
    if (-not $process.HasExited) {
        Stop-Process -Id $process.Id -Force
    }
}
