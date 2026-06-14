param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [Parameter(Mandatory = $true)]
    [string]$OutJson,

    [string]$ExpectedDisplaySizeName = "",
    [int]$ExpectedSizeHalfPoints = 21
)

$ErrorActionPreference = "Stop"

function Get-FileSha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Join-Chars([int[]]$Codes) {
    return -join ($Codes | ForEach-Object { [char]$_ })
}

function Get-NamedSize([double]$Points) {
    $map = @{
        "42" = (Join-Chars @(0x521D, 0x53F7)); "36" = (Join-Chars @(0x5C0F, 0x521D));
        "26" = (Join-Chars @(0x4E00, 0x53F7)); "24" = (Join-Chars @(0x5C0F, 0x4E00));
        "22" = (Join-Chars @(0x4E8C, 0x53F7)); "18" = (Join-Chars @(0x5C0F, 0x4E8C));
        "16" = (Join-Chars @(0x4E09, 0x53F7)); "15" = (Join-Chars @(0x5C0F, 0x4E09));
        "14" = (Join-Chars @(0x56DB, 0x53F7)); "12" = (Join-Chars @(0x5C0F, 0x56DB));
        "10.5" = (Join-Chars @(0x4E94, 0x53F7)); "9" = (Join-Chars @(0x5C0F, 0x4E94));
        "7.5" = (Join-Chars @(0x516D, 0x53F7)); "6.5" = (Join-Chars @(0x5C0F, 0x516D));
        "5.5" = (Join-Chars @(0x4E03, 0x53F7)); "5" = (Join-Chars @(0x516B, 0x53F7))
    }
    $key = ([double]$Points).ToString("0.###", [Globalization.CultureInfo]::InvariantCulture)
    if ($map.ContainsKey($key)) { return $map[$key] }
    return $key
}

function Read-FontProperty($Font, [string]$Name) {
    try {
        return [string]($Font.$Name)
    } catch {
        return ""
    }
}

function Get-ReferenceEntryIndex($Paragraph, [string]$Text) {
    if ($Text -match "^\s*\[(\d{1,3})\]") {
        return [int]$Matches[1]
    }
    if ($Text -match "^\s*(\d{1,3})(?:[\.．、])(?=\s|[^\d])") {
        return [int]$Matches[1]
    }
    $listString = ""
    try {
        $listString = [string]$Paragraph.Range.ListFormat.ListString
    } catch {
        $listString = ""
    }
    if ($listString -match "\d{1,3}") {
        return [int]$Matches[0]
    }
    return $null
}

$resolvedDocx = (Resolve-Path -LiteralPath $DocxPath).Path
$ReferenceHeading = Join-Chars @(0x53C2, 0x8003, 0x6587, 0x732E)
$AcknowledgementHeading = Join-Chars @(0x81F4, 0x8C22)
$AppendixHeading = Join-Chars @(0x9644, 0x5F55)
if ([string]::IsNullOrWhiteSpace($ExpectedDisplaySizeName)) {
    $ExpectedDisplaySizeName = Join-Chars @(0x4E94, 0x53F7)
}
$outPath = [System.IO.Path]::GetFullPath($OutJson)
$outDir = [System.IO.Path]::GetDirectoryName($outPath)
if ($outDir -and -not [System.IO.Directory]::Exists($outDir)) {
    [System.IO.Directory]::CreateDirectory($outDir) | Out-Null
}

$app = $null
$appKind = $null
foreach ($progId in @("KWPS.Application", "Word.Application")) {
    try {
        $app = New-Object -ComObject $progId
        $appKind = $progId
        break
    } catch {
        $app = $null
    }
}
if ($null -eq $app) {
    throw "Neither WPS (KWPS.Application) nor Word.Application COM automation is available."
}

$doc = $null
try {
    $app.Visible = $false
    $doc = $app.Documents.Open($resolvedDocx, $false, $true)
    $entries = New-Object System.Collections.Generic.List[object]
    $inReferences = $false
    $expectedPoints = [double]$ExpectedSizeHalfPoints / 2.0

    foreach ($paragraph in $doc.Paragraphs) {
        $text = ([string]$paragraph.Range.Text).Trim()
        $normalized = ($text -replace "\s+", "").ToLowerInvariant()
        if (-not $inReferences) {
            if ($normalized -eq $ReferenceHeading -or $normalized -eq "references" -or $normalized -eq "bibliography") {
                $inReferences = $true
            }
            continue
        }
        if ($text.StartsWith($AcknowledgementHeading) -or $text.StartsWith($AppendixHeading) -or $text -match "^(acknowledgements?|appendix)\b") {
            break
        }
        $entryIndex = Get-ReferenceEntryIndex $paragraph $text
        if ($null -eq $entryIndex) {
            continue
        }

        $range = $paragraph.Range
        try { $range.Select() | Out-Null } catch {}
        $font = $range.Font
        $size = [double](Read-FontProperty $font "Size")
        $inferred = Get-NamedSize $size
        $entryVerdict = "pass"
        if ([math]::Abs($size - $expectedPoints) -gt 0.01 -or $inferred -ne $ExpectedDisplaySizeName) {
            $entryVerdict = "fail"
        }
        $entries.Add([ordered]@{
            entryIndex = $entryIndex
            rangeStart = [int]$range.Start
            rangeEnd = [int]$range.End
            selectedText = if ($text.Length -gt 160) { $text.Substring(0, 160) } else { $text }
            fontName = Read-FontProperty $font "Name"
            fontNameAscii = Read-FontProperty $font "NameAscii"
            fontNameFarEast = Read-FontProperty $font "NameFarEast"
            fontNameOther = Read-FontProperty $font "NameOther"
            fontSizePoints = $size
            inferredWpsDisplaySizeName = $inferred
            verdict = $entryVerdict
        }) | Out-Null
    }

    $overall = "pass"
    if ($entries.Count -eq 0) {
        $overall = "fail"
    } else {
        foreach ($entry in $entries) {
            if ($entry.verdict -ne "pass") { $overall = "fail" }
        }
    }

    $docxSha256 = Get-FileSha256 $resolvedDocx
    $officeVersion = [string]$app.Version
    $entryArray = @($entries.ToArray())

    $payload = [ordered]@{
        schema = "graduation-project-builder.wps-reference-entry-ui-font.v1"
        generator = "scripts/audit_wps_reference_entry_ui_font.ps1"
        officeApp = $appKind
        officeVersion = $officeVersion
        docxPath = $resolvedDocx
        docxSha256 = $docxSha256
        expectedWpsDisplaySizeName = $ExpectedDisplaySizeName
        expectedSizeHalfPoints = [string]$ExpectedSizeHalfPoints
        expectedSizePoints = $expectedPoints
        checkedEntryCount = $entries.Count
        verdict = $overall
        entries = $entryArray
    }
    $json = $payload | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($outPath, $json + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
} finally {
    if ($null -ne $doc) { $doc.Close($false) | Out-Null }
    if ($null -ne $app) { $app.Quit() | Out-Null }
}
