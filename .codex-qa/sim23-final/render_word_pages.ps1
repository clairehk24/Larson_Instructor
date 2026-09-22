param(
    [string]$Only = '',
    [int]$StartPage = 1
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ClipboardMetafile {
    [DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr hWndNewOwner);
    [DllImport("user32.dll")] public static extern bool CloseClipboard();
    [DllImport("user32.dll")] public static extern IntPtr GetClipboardData(uint uFormat);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, string lpszFile);
    [DllImport("gdi32.dll")] public static extern bool DeleteEnhMetaFile(IntPtr hemf);
}
'@

$workspace = (Get-Location).Path
$wordPath = 'C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE'
$jobs = @(
    @{ Name = 'sp'; Input = 'assets\downloads\simulation-23\instructions-for-sp-1-and-2.docx' },
    @{ Name = 'proctors'; Input = 'assets\downloads\simulation-23\instructions-for-proctors.docx' },
    @{ Name = 'rubric1'; Input = 'assets\downloads\simulation-23\rubric-ehs-station-1.docx' },
    @{ Name = 'rubric2'; Input = 'assets\downloads\simulation-23\rubric-ehs-station-2.docx' }
)
if ($Only) {
    $jobs = @($jobs | Where-Object { $_.Name -eq $Only })
}

$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($job in $jobs) {
        $inputPath = (Resolve-Path $job.Input).Path
        $outputDir = Join-Path $workspace ('.codex-qa\sim23-final\' + $job.Name)
        New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
        $document = $word.Documents.Open($inputPath, $false, $true)
        try {
            $pageCount = $document.ComputeStatistics(2)
            Write-Output ("{0}:PAGES={1}" -f $job.Name, $pageCount)
            for ($page = $StartPage; $page -le $pageCount; $page++) {
                $start = $document.GoTo(1, 1, $page).Start
                if ($page -lt $pageCount) {
                    $end = $document.GoTo(1, 1, $page + 1).Start - 1
                } else {
                    $end = $document.Content.End - 1
                }
                $range = $document.Range($start, $end)
                [System.Windows.Forms.Clipboard]::Clear()
                $range.CopyAsPicture()
                $image = $null
                for ($attempt = 0; $attempt -lt 10 -and $null -eq $image; $attempt++) {
                    Start-Sleep -Milliseconds 500
                    $image = [System.Windows.Forms.Clipboard]::GetImage()
                }
                if ($null -eq $image) {
                    $emfPath = Join-Path $outputDir ("page-{0}.emf" -f $page)
                    if ([ClipboardMetafile]::OpenClipboard([IntPtr]::Zero)) {
                        try {
                            $sourceHandle = [ClipboardMetafile]::GetClipboardData(14)
                            if ($sourceHandle -ne [IntPtr]::Zero) {
                                $copyHandle = [ClipboardMetafile]::CopyEnhMetaFile($sourceHandle, $emfPath)
                                if ($copyHandle -ne [IntPtr]::Zero) {
                                    [ClipboardMetafile]::DeleteEnhMetaFile($copyHandle) | Out-Null
                                }
                            }
                        }
                        finally {
                            [ClipboardMetafile]::CloseClipboard() | Out-Null
                        }
                    }
                    if (Test-Path -LiteralPath $emfPath) {
                        $metafile = New-Object System.Drawing.Imaging.Metafile($emfPath)
                        $image = New-Object System.Drawing.Bitmap($metafile.Width, $metafile.Height)
                        $graphics = [System.Drawing.Graphics]::FromImage($image)
                        $graphics.Clear([System.Drawing.Color]::White)
                        $graphics.DrawImage($metafile, 0, 0, $image.Width, $image.Height)
                        $graphics.Dispose()
                        $metafile.Dispose()
                        Remove-Item -LiteralPath $emfPath -Force
                    }
                }
                if ($null -eq $image) {
                    $formats = [System.Windows.Forms.Clipboard]::GetDataObject().GetFormats() -join ', '
                    throw "Word did not render $($job.Name) page $page. Clipboard formats: $formats"
                }
                $outputPath = Join-Path $outputDir ("page-{0}.png" -f $page)
                $image.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $image.Dispose()
                Write-Output $outputPath
            }
        }
        finally {
            $document.Close(0)
        }
    }
    $word.Quit(0)
}
finally {
    try { $word.Quit(0) } catch { }
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
}
