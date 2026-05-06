param(
    [string]$WorkbookPath = ".\Database (5 mai 2026)\WOAH-DOI-database-v4_(CopyForDIAD).xlsx",
    [string]$OutputDir = ".\downloads-test",
    [int]$StartRow = 3295,
    [int]$EndRow = 3738
)

$ErrorActionPreference = "Stop"

function Clean-FileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return "untitled"
    }

    foreach ($ch in [IO.Path]::GetInvalidFileNameChars()) {
        $Name = $Name.Replace([string]$ch, "_")
    }

    $Name = $Name.Trim()
    if ($Name.Length -gt 120) {
        $Name = $Name.Substring(0, 120).Trim()
    }

    return $Name
}

function Resolve-Link {
    param(
        [string]$BaseUrl,
        [string]$Href
    )

    $decoded = [System.Net.WebUtility]::HtmlDecode($Href)
    if ($decoded -match "^https?://") {
        return $decoded
    }

    return ([Uri]::new([Uri]$BaseUrl, $decoded)).AbsoluteUri
}

function Test-PdfFile {
    param([string]$Path)

    $file = Get-Item -LiteralPath $Path
    if ($file.Length -lt 1024) {
        return [pscustomobject]@{
            IsValid = $false
            Issue = "File is too small to be a valid PDF"
        }
    }

    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $headerBytes = New-Object byte[] 5
        [void]$stream.Read($headerBytes, 0, $headerBytes.Length)
        $header = [Text.Encoding]::ASCII.GetString($headerBytes)

        if ($header -ne "%PDF-") {
            return [pscustomobject]@{
                IsValid = $false
                Issue = "File does not start with the PDF signature"
            }
        }

        $tailLength = [Math]::Min(4096, [int]$stream.Length)
        $stream.Seek(-1 * $tailLength, [IO.SeekOrigin]::End) | Out-Null
        $tailBytes = New-Object byte[] $tailLength
        [void]$stream.Read($tailBytes, 0, $tailBytes.Length)
        $tail = [Text.Encoding]::ASCII.GetString($tailBytes)

        if ($tail -notmatch "%%EOF") {
            return [pscustomobject]@{
                IsValid = $false
                Issue = "File does not contain a PDF EOF marker near the end"
            }
        }

        return [pscustomobject]@{
            IsValid = $true
            Issue = ""
        }
    }
    finally {
        $stream.Dispose()
    }
}

$resolvedWorkbook = Resolve-Path -LiteralPath $WorkbookPath
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$cleanPdfDir = Join-Path $OutputDir "clean-pdfs"
$tempDir = Join-Path $OutputDir "_temp"
$reportPath = Join-Path $OutputDir "download_report.csv"

New-Item -ItemType Directory -Force -Path $cleanPdfDir | Out-Null
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

$report = New-Object System.Collections.Generic.List[object]

$excel = $null
$workbook = $null
$openedHere = $false

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
    }

    foreach ($book in $excel.Workbooks) {
        if ($book.FullName -eq $resolvedWorkbook.Path) {
            $workbook = $book
            break
        }
    }

    if ($null -eq $workbook) {
        $workbook = $excel.Workbooks.Open($resolvedWorkbook.Path)
        $openedHere = $true
    }

    $worksheet = $workbook.Worksheets.Item("DATABASE")
    $usedRange = $worksheet.UsedRange
    $data = $usedRange.Value2

    $lastRow = [Math]::Min($usedRange.Rows.Count, $EndRow)

    for ($row = $StartRow; $row -le $lastRow; $row++) {
        $rank = [string]$data[$row, 1]
        $title = [string]$data[$row, 7]
        $doi = [string]$data[$row, 12]
        $url = [string]$data[$row, 13]

        if ([string]::IsNullOrWhiteSpace($url)) {
            $url = $doi
        }

        if ([string]::IsNullOrWhiteSpace($url)) {
            Write-Warning "Row $row has no URL or DOI. Skipping."
            $report.Add([pscustomobject]@{
                Row = $row
                Rank = $rank
                Title = $title
                DOI = $doi
                RecordUrl = $url
                DownloadUrl = ""
                OutputFile = ""
                Status = "MissingUrl"
                Issue = "No URL or DOI in row"
                Bytes = 0
            })
            continue
        }

        Write-Host "Row ${row}: opening $url"

        try {
            $page = Invoke-WebRequest -Uri $url -MaximumRedirection 5 -TimeoutSec 30 -UseBasicParsing

            $pdfLinks = $page.Links |
                Where-Object { $_.href -match "digidoc\.xhtml" -and $_.href -match "downloadAttachment" } |
                Select-Object -ExpandProperty href -Unique

            if (-not $pdfLinks) {
                Write-Warning "Row ${row}: no PDF download link found."
                $report.Add([pscustomobject]@{
                    Row = $row
                    Rank = $rank
                    Title = $title
                    DOI = $doi
                    RecordUrl = $url
                    DownloadUrl = ""
                    OutputFile = ""
                    Status = "NoPdfLink"
                    Issue = "No digidoc PDF download link found on portal page"
                    Bytes = 0
                })
                continue
            }

            $downloadIndex = 1
            foreach ($link in $pdfLinks) {
                $downloadUrl = Resolve-Link -BaseUrl $url -Href $link

                $fileStem = Clean-FileName "$rank - $title"
                if ($pdfLinks.Count -gt 1) {
                    $fileStem = "$fileStem - $downloadIndex"
                }

                $targetFile = Join-Path $cleanPdfDir "$fileStem.pdf"
                $tempFile = Join-Path $tempDir "$([guid]::NewGuid()).pdf"

                Write-Host "Row ${row}: downloading and validating $targetFile"
                Invoke-WebRequest -Uri $downloadUrl -OutFile $tempFile -MaximumRedirection 5 -TimeoutSec 60 -UseBasicParsing

                $validation = Test-PdfFile -Path $tempFile
                $bytes = (Get-Item -LiteralPath $tempFile).Length

                if ($validation.IsValid) {
                    Move-Item -LiteralPath $tempFile -Destination $targetFile -Force
                    $report.Add([pscustomobject]@{
                        Row = $row
                        Rank = $rank
                        Title = $title
                        DOI = $doi
                        RecordUrl = $url
                        DownloadUrl = $downloadUrl
                        OutputFile = $targetFile
                        Status = "Clean"
                        Issue = ""
                        Bytes = $bytes
                    })
                    Write-Host "Row ${row}: clean PDF saved"
                }
                else {
                    Remove-Item -LiteralPath $tempFile -Force
                    $report.Add([pscustomobject]@{
                        Row = $row
                        Rank = $rank
                        Title = $title
                        DOI = $doi
                        RecordUrl = $url
                        DownloadUrl = $downloadUrl
                        OutputFile = ""
                        Status = "CorruptedPdf"
                        Issue = $validation.Issue
                        Bytes = $bytes
                    })
                    Write-Warning "Row ${row}: corrupted PDF discarded - $($validation.Issue)"
                }

                $downloadIndex++
                Start-Sleep -Milliseconds 500
            }
        }
        catch {
            Write-Warning "Row $row failed: $($_.Exception.Message)"
            $report.Add([pscustomobject]@{
                Row = $row
                Rank = $rank
                Title = $title
                DOI = $doi
                RecordUrl = $url
                DownloadUrl = ""
                OutputFile = ""
                Status = "Failed"
                Issue = $_.Exception.Message
                Bytes = 0
            })
        }
    }
}
finally {
    if ($report.Count -gt 0) {
        $report | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report written to $reportPath"
        Write-Host "Clean PDFs are in $cleanPdfDir"
    }

    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }

    if ($openedHere -and $null -ne $workbook) {
        $workbook.Close($false)
    }

    if ($null -ne $excel -and $excel.Workbooks.Count -eq 0) {
        $excel.Quit()
    }
}
