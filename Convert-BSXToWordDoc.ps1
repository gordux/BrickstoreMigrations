#Requires -Version 7.0
<#
.SYNOPSIS
    Converts a BrickStore BSX file into a Word (.docx) document with BrickLink images.

.DESCRIPTION
    Pure PowerShell 7 / .NET only — no Node.js, no external tools, no modules required.
    Runs in a single pass:
      1. Parses the BSX XML
      2. Optionally filters to specific item types
      3. Downloads & caches product images from BrickLink
      4. Builds a valid .docx (ZIP of XML) entirely in memory
      5. Two items per page, each with image + full metadata card

.PARAMETER BSXFile
    Path to the .bsx file to process.

.PARAMETER InventoryName
    Name printed in the document title (e.g. "BrickStore", "My LEGO Collection").
    Required.

.PARAMETER OutputDocx
    Path for the output .docx. Defaults to same folder/name as the BSX file.

.PARAMETER ImageCacheDir
    Folder to cache downloaded images. Defaults to the system temp folder.

.PARAMETER ItemTypes
    One or more item types to include. Omit to include all types.
    Accepted values (case-insensitive, friendly names or single-letter codes):
        Part  / Parts  / P
        Set   / Sets   / S
        Minifig / Minifigure / Minifigures / M
        Book  / Books  / B
        Gear  / G
        Catalog / Catalogs / C
        Instruction / Instructions / I
        OriginalBox / OriginalBoxes / O
    Combine multiple:  -ItemTypes Set,Minifig,Part

.EXAMPLE
    .\Convert-BSXToWordDoc.ps1 -BSXFile "C:\LEGO\MySet.bsx"

.EXAMPLE
    .\Convert-BSXToWordDoc.ps1 -BSXFile "C:\LEGO\MySet.bsx" -OutputDocx "C:\Output\MySet.docx"

.EXAMPLE
    .\Convert-BSXToWordDoc.ps1 -BSXFile "C:\LEGO\MySet.bsx" -ItemTypes Set,Minifig

.EXAMPLE
    .\Convert-BSXToWordDoc.ps1 -BSXFile "C:\LEGO\MySet.bsx" -ItemTypes P

.EXAMPLE
    .\Convert-BSXToWordDoc.ps1 -BSXFile "C:\LEGO\MySet.bsx" -ItemTypes Parts,Sets,Gear
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$BSXFile,

    [Parameter(Mandatory = $true)]
    [string]$InventoryName,

    [Parameter(Mandatory = $false)]
    [string]$OutputDocx = "",

    [Parameter(Mandatory = $false)]
    [string]$ImageCacheDir = "",

    # One or more item types: friendly names, plural forms, or raw BSX codes
    # e.g.  -ItemTypes Set,Minifig,Part   or   -ItemTypes S,M,P
    [Parameter(Mandatory = $false)]
    [string[]]$ItemTypes = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ══════════════════════════════════════════════════════════════
# 1. Paths
# ══════════════════════════════════════════════════════════════

$BSXFile = (Resolve-Path $BSXFile).Path

if (-not $OutputDocx) {
    $base       = [System.IO.Path]::GetFileNameWithoutExtension($BSXFile)
    $OutputDocx = Join-Path (Split-Path $BSXFile -Parent) "$base.docx"
}

# Use [System.IO.Path]::GetTempPath() for cross-platform compatibility (Win/Mac/Linux)
if (-not $ImageCacheDir) {
    $ImageCacheDir = Join-Path ([System.IO.Path]::GetTempPath()) "BSX_ImageCache"
}
if (-not (Test-Path $ImageCacheDir)) {
    New-Item -ItemType Directory -Path $ImageCacheDir | Out-Null
}

# ── Resolve -ItemTypes to canonical single-letter BSX codes ──
$typeAliasMap = @{
    'P'='P';'PART'='P';'PARTS'='P'
    'S'='S';'SET'='S';'SETS'='S'
    'M'='M';'MINIFIG'='M';'MINIFIGS'='M';'MINIFIGURE'='M';'MINIFIGURES'='M'
    'B'='B';'BOOK'='B';'BOOKS'='B'
    'G'='G';'GEAR'='G'
    'C'='C';'CATALOG'='C';'CATALOGS'='C'
    'I'='I';'INSTRUCTION'='I';'INSTRUCTIONS'='I'
    'O'='O';'ORIGINALBOX'='O';'ORIGINALBOXES'='O'
}

$typeDisplayNames = @{
    'P'='Parts';'S'='Sets';'M'='Minifigures'
    'B'='Books';'G'='Gear';'C'='Catalogs'
    'I'='Instructions';'O'='Original Boxes'
}

$filterCodes = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
foreach ($entry in $ItemTypes) {
    $key = $entry.Trim().ToUpper()
    if ($typeAliasMap.ContainsKey($key)) {
        [void]$filterCodes.Add($typeAliasMap[$key])
    } else {
        Write-Warning "Unknown item type '$entry' — ignored. Valid: Part,Set,Minifig,Book,Gear,Catalog,Instruction,OriginalBox (or P,S,M,B,G,C,I,O)"
    }
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════╗"
Write-Host "║     BSX -> Word Document Converter       ║"
Write-Host "╚══════════════════════════════════════════╝"
Write-Host "  Input : $BSXFile"
Write-Host "  Output: $OutputDocx"
if ($filterCodes.Count -gt 0) {
    $friendlyList = ($filterCodes | ForEach-Object { $typeDisplayNames[$_] }) -join ", "
    Write-Host "  Filter : $friendlyList"
} else {
    Write-Host "  Filter : All item types"
}
Write-Host ""

# ══════════════════════════════════════════════════════════════
# 2. Parse BSX (XML)
# ══════════════════════════════════════════════════════════════

Write-Host "[1/4] Parsing BSX file..."

# Load XML via XmlDocument for reliable UTF-8 handling in PS7
$xmlDoc = [System.Xml.XmlDocument]::new()
$xmlDoc.Load($BSXFile)

# Helper: safely read a child element's inner text without dot-notation on XmlNode,
# which throws under Set-StrictMode -Version Latest in PS7 when the node is absent.
function Get-NodeText ([System.Xml.XmlNode]$parent, [string]$childName) {
    $child = $parent.SelectSingleNode($childName)
    if ($null -ne $child) { return [string]$child.InnerText } else { return $null }
}

# Use PSCustomObject throughout — dot-notation works correctly under Set-StrictMode
$allItems = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($node in $xmlDoc.SelectNodes('//BrickStoreXML/Inventory/Item')) {

    $t  = (Get-NodeText $node 'ItemTypeID')  ?? 'P'
    $id = (Get-NodeText $node 'ItemID')    ?? ''
    $co = (Get-NodeText $node 'ColorID')   ?? '0'

    $rawCond = Get-NodeText $node 'Condition'
    $cond = switch ($rawCond) {
        'N'     { 'New' }
        'U'     { 'Used' }
        default { $rawCond ?? '' }
    }

    $imgUrl = switch ($t.ToUpper()) {
        'S'     { "https://img.bricklink.com/ItemImage/SN/0/$id.png" }
        'M'     { "https://img.bricklink.com/ItemImage/MN/0/$id.png" }
        'B'     { "https://img.bricklink.com/ItemImage/BN/0/$id.png" }
        'G'     { "https://img.bricklink.com/ItemImage/GN/0/$id.png" }
        default { "https://img.bricklink.com/ItemImage/PN/$co/$id.png" }
    }

    $itemName = (Get-NodeText $node 'ItemName') ?? $id

    $allItems.Add([PSCustomObject]@{
        ItemType     = $t
        ItemID       = $id
        ColorID      = $co
        ColorName    = (Get-NodeText $node 'ColorName')    ?? ''
        ItemName     = $itemName
        CategoryName = (Get-NodeText $node 'CategoryName') ?? ''
        Condition    = $cond
        Qty          = (Get-NodeText $node 'Qty')          ?? '1'
        Price        = (Get-NodeText $node 'Price')        ?? ''
        Comments     = (Get-NodeText $node 'Comments')     ?? ''
        Remarks      = (Get-NodeText $node 'Remarks')      ?? ''
        ImageUrl     = $imgUrl
        ImagePath    = ''
    })
}

Write-Host "      Found $($allItems.Count) items in BSX file."

# ── Apply item-type filter ────────────────────────────────────
$items = [System.Collections.Generic.List[PSCustomObject]]::new()
if ($filterCodes.Count -gt 0) {
    foreach ($item in $allItems) {
        if ($filterCodes.Contains($item.ItemType.ToUpper())) {
            $items.Add($item)
        }
    }
    Write-Host "      After filter: $($items.Count) of $($allItems.Count) items included."
    if ($items.Count -eq 0) {
        Write-Warning "No items matched the requested type(s). The document will be empty."
    }
} else {
    foreach ($item in $allItems) { $items.Add($item) }
}

# Sort by theme (CategoryName) then by item name — blank theme sorts to top as "(No Theme)"
$items = [System.Collections.Generic.List[PSCustomObject]](
    $items | Sort-Object -Property @(
        { if ($_.CategoryName) { $_.CategoryName } else { '(No Theme)' } },
        { $_.ItemName }
    )
)

# ══════════════════════════════════════════════════════════════
# 3. Download / cache images
# ══════════════════════════════════════════════════════════════

Write-Host "[2/4] Downloading and resizing images from BrickLink..."

# Target size for embedded images — 210px matches the enlarged card image column
$resizeTarget = 210

# Load System.Drawing for image resizing (available on Windows with PS7)
try {
    Add-Type -AssemblyName System.Drawing
    $drawingAvailable = $true
} catch {
    Write-Warning "System.Drawing not available — images will be embedded at original size."
    $drawingAvailable = $false
}

# Resize a PNG/JPEG to fit within $maxPx x $maxPx, preserving aspect ratio.
# Returns the resized image as a PNG byte array.
function Resize-ImageBytes ([byte[]]$srcBytes, [int]$maxPx) {
    $srcStream  = [System.IO.MemoryStream]::new($srcBytes)
    $srcBitmap  = [System.Drawing.Bitmap]::new($srcStream)

    $srcW = $srcBitmap.Width
    $srcH = $srcBitmap.Height

    # Only downscale — never upscale small images
    if ($srcW -le $maxPx -and $srcH -le $maxPx) {
        $srcBitmap.Dispose()
        $srcStream.Dispose()
        return $srcBytes
    }

    $scale  = [math]::Min($maxPx / $srcW, $maxPx / $srcH)
    $dstW   = [int]($srcW * $scale)
    $dstH   = [int]($srcH * $scale)

    $dstBitmap = [System.Drawing.Bitmap]::new($dstW, $dstH)
    $graphics  = [System.Drawing.Graphics]::FromImage($dstBitmap)
    $graphics.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graphics.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $graphics.DrawImage($srcBitmap, 0, 0, $dstW, $dstH)

    $outStream = [System.IO.MemoryStream]::new()
    $dstBitmap.Save($outStream, [System.Drawing.Imaging.ImageFormat]::Png)

    $graphics.Dispose()
    $dstBitmap.Dispose()
    $srcBitmap.Dispose()
    $srcStream.Dispose()

    return $outStream.ToArray()
}

# Minimal grey placeholder PNG (1x1 pixel, base64)
$placeholderBytes = [Convert]::FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAAIAAAACAAQAAAADrRVxmAAAADklEQVR4nGNgGAWDHQAAAiAAAR" +
    "HzJkAAAAAASUVORK5CYII="
)
$placeholderPath = Join-Path $ImageCacheDir "placeholder.png"
[System.IO.File]::WriteAllBytes($placeholderPath, $placeholderBytes)

# PS7: Use Invoke-WebRequest instead of the deprecated System.Net.WebClient
$iwrParams = @{
    UseBasicParsing = $true
    Headers         = @{ 'User-Agent' = 'Mozilla/5.0 BSX-DocConverter/3.0' }
    TimeoutSec      = 15
    ErrorAction     = 'Stop'
}

$resizedCacheDir = Join-Path $ImageCacheDir "resized_${resizeTarget}px"
if (-not (Test-Path $resizedCacheDir)) {
    New-Item -ItemType Directory -Path $resizedCacheDir | Out-Null
}

$done = 0
foreach ($item in $items) {
    $done++
    $safeID      = $item.ItemID -replace '[\\/:*?"<>|]', '_'
    $cacheName   = "$($item.ItemType)_${safeID}_$($item.ColorID).png"
    $cachePath   = Join-Path $ImageCacheDir $cacheName
    $resizedPath = Join-Path $resizedCacheDir $cacheName

    # ── Fast path: resized cache hit — nothing to download or resize ──────────
    if ($drawingAvailable -and (Test-Path $resizedPath)) {
        $item.ImagePath = $resizedPath

    # ── Original cache hit — skip download, resize only if needed ─────────────
    } elseif (Test-Path $cachePath) {
        $isPlaceholderSrc = ((Get-Item $cachePath).Length -eq $placeholderBytes.Length)

        if ($drawingAvailable -and -not $isPlaceholderSrc) {
            try {
                $originalBytes = [System.IO.File]::ReadAllBytes($cachePath)
                $resizedBytes  = Resize-ImageBytes $originalBytes $resizeTarget
                [System.IO.File]::WriteAllBytes($resizedPath, $resizedBytes)
                $item.ImagePath = $resizedPath
            } catch {
                $item.ImagePath = $cachePath
            }
        } else {
            $item.ImagePath = $cachePath
        }

    # ── Cache miss — download, validate, then resize ──────────────────────────
    } else {
        try {
            Invoke-WebRequest @iwrParams -Uri $item.ImageUrl -OutFile $cachePath

            # Validate PNG/JPEG magic bytes — BrickLink returns HTML for missing items
            $magic = [System.IO.File]::ReadAllBytes($cachePath)
            $valid = ($magic.Length -gt 4) -and (
                ($magic[0] -eq 0x89 -and $magic[1] -eq 0x50) -or   # PNG
                ($magic[0] -eq 0xFF -and $magic[1] -eq 0xD8)         # JPEG
            )
            if (-not $valid) {
                Remove-Item $cachePath -Force
                Copy-Item $placeholderPath $cachePath
            }
        }
        catch {
            # Network error or 404 — use placeholder silently
            if (Test-Path $cachePath) { Remove-Item $cachePath -Force }
            Copy-Item $placeholderPath $cachePath
        }

        $isPlaceholderSrc = ((Get-Item $cachePath).Length -eq $placeholderBytes.Length)

        if ($drawingAvailable -and -not $isPlaceholderSrc) {
            try {
                $originalBytes = [System.IO.File]::ReadAllBytes($cachePath)
                $resizedBytes  = Resize-ImageBytes $originalBytes $resizeTarget
                [System.IO.File]::WriteAllBytes($resizedPath, $resizedBytes)
                $item.ImagePath = $resizedPath
            } catch {
                $item.ImagePath = $cachePath
            }
        } else {
            $item.ImagePath = $cachePath
        }
    }

    Write-Progress -Activity "Downloading and resizing images" `
                   -Status   "$done / $($items.Count)  [$($item.ItemID)]" `
                   -PercentComplete ([int](($done / $items.Count) * 100))
}

Write-Progress -Activity "Downloading and resizing images" -Completed
Write-Host "      Image download and resize complete."

# ══════════════════════════════════════════════════════════════
# 4. Build DOCX XML
# ══════════════════════════════════════════════════════════════

Write-Host "[3/4] Building document XML..."

function XmlEsc ([string]$s) {
    # Use SecurityElement.Escape for complete, reliable XML escaping in PS7
    return [System.Security.SecurityElement]::Escape($s)
}

function Get-DrawingXml ([string]$rId, [long]$cx, [long]$cy, [int]$id) {
@"
<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <wp:extent cx="$cx" cy="$cy"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:docPr id="$id" name="img$id"/>
  <wp:cNvGraphicFramePr>
    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
  </wp:cNvGraphicFramePr>
  <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:nvPicPr>
          <pic:cNvPr id="$id" name="img$id"/>
          <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
        </pic:nvPicPr>
        <pic:blipFill>
          <a:blip r:embed="$rId" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
          <a:stretch><a:fillRect/></a:stretch>
        </pic:blipFill>
        <pic:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="$cx" cy="$cy"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </pic:spPr>
      </pic:pic>
    </a:graphicData>
  </a:graphic>
</wp:inline></w:drawing>
"@
}

function Get-LabelValueXml ([string]$label, [string]$value) {
    $l = XmlEsc $label
    $v = XmlEsc $value
@"
<w:p>
  <w:pPr><w:spacing w:before="30" w:after="30"/></w:pPr>
  <w:r><w:rPr><w:b/><w:color w:val="444444"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr>
    <w:t xml:space="preserve">$l </w:t></w:r>
  <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
    <w:t>$v</w:t></w:r>
</w:p>
"@
}

$bodyParts   = [System.Collections.Generic.List[string]]::new()
$imageRelXml = [System.Collections.Generic.List[string]]::new()
$imageFiles  = [ordered]@{}   # rId -> PSCustomObject { FileName; Bytes; Ext }

$rIdN = 1
$imgN = 1

# Pre-register the placeholder as a single shared image entry so it is only
# embedded once in the ZIP regardless of how many items lack a real image.
$placeholderRId   = 'rIdPH'
$placeholderEmbed = [System.IO.File]::ReadAllBytes($placeholderPath)
$imageFiles[$placeholderRId] = [PSCustomObject]@{ FileName = 'img_placeholder.png'; Bytes = $placeholderEmbed; Ext = 'png' }
$imageRelXml.Add(
    "<Relationship Id=`"$placeholderRId`" " +
    "Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" " +
    "Target=`"media/img_placeholder.png`"/>"
)

$rIdN = 1
$imgN = 1

# Layout constants — must stay in sync with the table column widths below
# Page: 12240 DXA wide, 1080 DXA left+right margins → 10080 DXA content width
# Image column: 3120 DXA  (~2.17")  → 2,981,250 EMU  (constraining dimension)
# Detail column: 6060 DXA (~4.21")
# These are chosen so 4 cards fill a US Letter page with 0.75" margins.
$imgColDxa    = 3120
$detailColDxa = 6060
$tableWidthDxa = $imgColDxa + $detailColDxa   # 9180 DXA

# Convert image column width to EMU — this is the max square the image can fill
$EMU = [long]($imgColDxa * 914400 / 1440)     # 1,982,000 EMU ≈ 208px @ 96dpi

# ── Title block ───────────────────────────────────────────────
$dateStr = (Get-Date).ToString("MMMM d, yyyy")
$bodyParts.Add(@"
<w:p>
  <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="100"/></w:pPr>
  <w:r><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr>
    <w:t>$(XmlEsc $InventoryName) Inventory</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="400"/></w:pPr>
  <w:r><w:rPr><w:color w:val="777777"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
    <w:t>$(XmlEsc "$($items.Count) items  |  $dateStr")</w:t></w:r>
</w:p>
"@)

# ── Item cards ────────────────────────────────────────────────
$cardIdx      = 0   # total cards emitted (drives page breaks)
$cardsOnPage  = 0   # cards on the current page (resets after page break)
$currentTheme = $null

foreach ($item in $items) {
    $cardIdx++

    $theme = if ($item.CategoryName) { $item.CategoryName } else { '(No Theme)' }

    # ── Theme header ──────────────────────────────────────────
    if ($theme -ne $currentTheme) {
        # If we are mid-page when the theme changes, flush to a new page
        # (but not before the very first card)
        if ($null -ne $currentTheme -and $cardsOnPage -gt 0) {
            $bodyParts.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
            $cardsOnPage = 0
        }
        $currentTheme = $theme
        $bodyParts.Add(@"
<w:p>
  <w:pPr>
    <w:spacing w:before="0" w:after="160"/>
    <w:pBdr>
      <w:bottom w:val="single" w:sz="6" w:space="4" w:color="2E75B6"/>
    </w:pBdr>
  </w:pPr>
  <w:r><w:rPr><w:b/><w:color w:val="2E75B6"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
    <w:t>$(XmlEsc $theme)</w:t></w:r>
</w:p>
"@)
    }

    $imgId    = $imgN; $imgN++
    $rawBytes = [System.IO.File]::ReadAllBytes($item.ImagePath)
    $isPlaceholder = ($item.ImagePath -eq $placeholderPath)

    if ($isPlaceholder) {
        $rId = $placeholderRId
        $rIdN--
    } else {
        $rId    = "rId$rIdN"; $rIdN++
        $isJpeg = ($rawBytes[0] -eq 0xFF -and $rawBytes[1] -eq 0xD8)
        $ext    = if ($isJpeg) { 'jpeg' } else { 'png' }
        $fname  = "img${imgId}.$ext"
        $imageFiles[$rId] = [PSCustomObject]@{ FileName = $fname; Bytes = $rawBytes; Ext = $ext }
        $imageRelXml.Add(
            "<Relationship Id=`"$rId`" " +
            "Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" " +
            "Target=`"media/$fname`"/>"
        )
    }

    # Use InvariantCulture for decimal parsing — safe across all PS7 locales
    $priceStr = if ($item.Price) {
        try {
            '$' + [double]::Parse($item.Price, [System.Globalization.CultureInfo]::InvariantCulture).ToString('0.00')
        } catch {
            $item.Price
        }
    } else { '—' }

    $detailXml  = Get-LabelValueXml 'Item ID:'    $item.ItemID
    if ($item.Condition) { $detailXml += Get-LabelValueXml 'Condition:' $item.Condition }
    $detailXml += Get-LabelValueXml 'Quantity:'   $item.Qty
    $detailXml += Get-LabelValueXml 'Price:'      $priceStr
    if ($item.Comments)  { $detailXml += Get-LabelValueXml 'Comments:'  $item.Comments }

    # Compute display EMU preserving the image's true aspect ratio so it is
    # never stretched or squashed. Cap the long edge at $EMU (100px equivalent).
    $imgCx = $EMU
    $imgCy = $EMU
    if ($drawingAvailable -and -not $isPlaceholder) {
        try {
            $ms  = [System.IO.MemoryStream]::new($rawBytes)
            $bmp = [System.Drawing.Bitmap]::new($ms)
            $pw  = $bmp.Width
            $ph  = $bmp.Height
            $bmp.Dispose(); $ms.Dispose()
            if ($pw -gt 0 -and $ph -gt 0) {
                $scale = [math]::Min($EMU / $pw, $EMU / $ph)
                $imgCx = [long]($pw * $scale)
                $imgCy = [long]($ph * $scale)
            }
        } catch { <# fall back to square EMU #> }
    }

    $bodyParts.Add(@"
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="$tableWidthDxa" w:type="dxa"/>
    <w:tblBorders>
      <w:top     w:val="single" w:sz="10" w:space="0" w:color="2E75B6"/>
      <w:left    w:val="single" w:sz="4"  w:space="0" w:color="CCCCCC"/>
      <w:bottom  w:val="single" w:sz="4"  w:space="0" w:color="CCCCCC"/>
      <w:right   w:val="single" w:sz="4"  w:space="0" w:color="CCCCCC"/>
      <w:insideH w:val="none"/>
      <w:insideV w:val="none"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top    w:w="80" w:type="dxa"/>
      <w:left   w:w="120" w:type="dxa"/>
      <w:bottom w:w="80" w:type="dxa"/>
      <w:right  w:w="120" w:type="dxa"/>
    </w:tblCellMar>
    <w:tblLook w:val="0000"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="$imgColDxa"/>
    <w:gridCol w:w="$detailColDxa"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="$imgColDxa" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
        </w:tcBorders>
        <w:vAlign w:val="center"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/></w:pPr>
        <w:r>$(Get-DrawingXml $rId $imgCx $imgCy $imgId)</w:r>
      </w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="$detailColDxa" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
        </w:tcBorders>
        <w:vAlign w:val="top"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:spacing w:before="0" w:after="60"/></w:pPr>
        <w:r><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
          <w:t>$(XmlEsc $item.ItemName)</w:t></w:r>
      </w:p>
      $detailXml
    </w:tc>
  </w:tr>
</w:tbl>
<w:p><w:pPr><w:spacing w:before="0" w:after="140"/></w:pPr></w:p>
"@)

    $cardsOnPage++
    if (($cardsOnPage % 4 -eq 0) -and ($cardIdx -lt $items.Count)) {
        $bodyParts.Add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        $cardsOnPage = 0
    }
}

# ══════════════════════════════════════════════════════════════
# 5. Compose XML file strings
# ══════════════════════════════════════════════════════════════

$uniqueExts = ($imageFiles.Values | ForEach-Object { $_.Ext } | Sort-Object -Unique)
$ctDefaults = $uniqueExts | ForEach-Object {
    $mime = if ($_ -eq 'jpeg') { 'image/jpeg' } else { 'image/png' }
    "  <Default Extension=`"$_`" ContentType=`"$mime`"/>"
}

$xmlContentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
$($ctDefaults -join "`n")
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"@

$xmlRootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>
"@

$xmlDocRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSt"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rIdSe"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
  $($imageRelXml -join "`n  ")
</Relationships>
"@

$xmlSettings = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
"@

$xmlStyles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="20"/><w:szCs w:val="20"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr>
      <w:tblBorders>
        <w:top    w:val="single" w:sz="4" w:color="auto"/>
        <w:left   w:val="single" w:sz="4" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:color="auto"/>
        <w:right  w:val="single" w:sz="4" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:color="auto"/>
        <w:insideV w:val="single" w:sz="4" w:color="auto"/>
      </w:tblBorders>
    </w:tblPr>
  </w:style>
</w:styles>
"@

$bodyInner   = $bodyParts -join "`n"
$xmlDocument = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  mc:Ignorable="w14">
  <w:body>
    $bodyInner
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"
               w:header="360" w:footer="360" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

# ══════════════════════════════════════════════════════════════
# 6. Write .docx (ZIP)
# ══════════════════════════════════════════════════════════════

Write-Host "[4/4] Writing .docx..."

# System.IO.Compression is part of the .NET runtime in PS7 — no Add-Type needed
# Including it anyway is harmless, but omitting avoids spurious assembly warnings
$null = [System.IO.Compression.ZipArchive]  # ensure assembly is loaded

if (Test-Path $OutputDocx) { Remove-Item $OutputDocx -Force }

# UTF-8 without BOM — required for valid XML inside the ZIP
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Add-ZipEntry {
    param(
        [System.IO.Compression.ZipArchive] $Archive,
        [string] $EntryName,
        [byte[]] $Data
    )
    $entry  = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::SmallestSize)
    $stream = $entry.Open()
    $stream.Write($Data, 0, $Data.Length)
    $stream.Dispose()
}

$fs  = [System.IO.FileStream]::new(
    $OutputDocx,
    [System.IO.FileMode]::Create,
    [System.IO.FileAccess]::Write
)
$zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)

try {
    Add-ZipEntry -Archive $zip -EntryName '[Content_Types].xml'          -Data ($utf8NoBom.GetBytes($xmlContentTypes))
    Add-ZipEntry -Archive $zip -EntryName '_rels/.rels'                  -Data ($utf8NoBom.GetBytes($xmlRootRels))
    Add-ZipEntry -Archive $zip -EntryName 'word/document.xml'            -Data ($utf8NoBom.GetBytes($xmlDocument))
    Add-ZipEntry -Archive $zip -EntryName 'word/_rels/document.xml.rels' -Data ($utf8NoBom.GetBytes($xmlDocRels))
    Add-ZipEntry -Archive $zip -EntryName 'word/settings.xml'            -Data ($utf8NoBom.GetBytes($xmlSettings))
    Add-ZipEntry -Archive $zip -EntryName 'word/styles.xml'              -Data ($utf8NoBom.GetBytes($xmlStyles))

    foreach ($rId in $imageFiles.Keys) {
        $info = $imageFiles[$rId]
        Add-ZipEntry -Archive $zip -EntryName "word/media/$($info.FileName)" -Data $info.Bytes
    }
}
finally {
    $zip.Dispose()
    $fs.Dispose()
}

# ══════════════════════════════════════════════════════════════
# 7. Summary
# ══════════════════════════════════════════════════════════════

$sizeKB = [math]::Round((Get-Item $OutputDocx).Length / 1KB, 1)
$pages  = [math]::Ceiling($items.Count / 4)

$typeBreakdown = $items |
    Group-Object -Property ItemType |
    Sort-Object Name |
    ForEach-Object {
        $code  = $_.Name.ToUpper()
        $label = if ($typeDisplayNames.ContainsKey($code)) { $typeDisplayNames[$code] } else { $code }
        "    $label : $($_.Count)"
    }

Write-Host ""
Write-Host "╔══════════════════════════════════════════╗"
Write-Host "║  Done!                                   ║"
Write-Host "╚══════════════════════════════════════════╝"
Write-Host "  Items  : $($items.Count)"
if ($typeBreakdown) {
    Write-Host "  By type:"
    $typeBreakdown | ForEach-Object { Write-Host $_ }
}
Write-Host "  Pages  : $pages"
Write-Host "  Output : $OutputDocx"
Write-Host "  Size   : $sizeKB KB"
Write-Host ""
