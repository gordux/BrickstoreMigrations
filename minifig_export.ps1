<#
.SYNOPSIS
    Converts a BrickStore BSX file into a text file that can be used to view minifigures id in a table..

.DESCRIPTION
    Pure PowerShell 7 / .NET only — no Node.js, no external tools, no modules required.
    Runs in a single pass:
      1. Parses the BSX XML
      2. Outputs only minifigures with known series prefixes

.PARAMETER BSXFile
    Path to the .bsx file to process.

.PARAMETER OutputDocx
    Path for the output .docx. Defaults to same folder/name as the BSX file.

.EXAMPLE
    .\minifig_export.ps1 -BSXFile "C:\LEGO\MySet.bsx"

.EXAMPLE
    .\minifig_export.ps1 -BSXFile "C:\LEGO\MySet.bsx" -OutputDocx "C:\Output\MySet.txt"
#>

param (
    # Path to the input BSX file (required)
    [Parameter(Mandatory = $true)]
    [string]$BsxPath,

    # Path for the output text file; defaults to the same directory and base
    # name as the BSX file if not provided
    [Parameter(Mandatory = $false)]
    [string]$OutputDocx = ""
)

# If no output path was specified, derive one from the input file path
if (-not $OutputDocx) {
    $base       = [System.IO.Path]::GetFileNameWithoutExtension($BsxPath)
    $OutputDocx = Join-Path (Split-Path $BsxPath -Parent) "$base.txt"
}

# Load the BSX file as XML so its nodes can be queried directly
[xml]$bsx = Get-Content -Path $BsxPath

# Series lookup table: maps a regex prefix pattern (matched against a minifig
# ItemID) to a human-readable series name. Add new entries here as needed.
$SeriesLookup = @{
    "^col"        = "Collectible Minifigures"
    "^sw"         = "Star Wars"
    "^hp"         = "Harry Potter"
    "^njo"        = "Ninjago"
    "^sh"         = "Super Heroes"
    "^lor"        = "Lord of the Rings"
    "^hob"        = "The Hobbit"
    "^pir"        = "Pirates"
    "^cas"        = "Castle"
    "^adv"        = "Adventurers"
    "^fig"        = "Generic / Unclassified"
    "^cty"        = "City"
    "^frnd"       = "Friends"
    "^pi"         = "Classic Pirates"
    "^poc"        = "Pirates of the Carribean"
}

# Returns the series name for a given minifig ID by testing it against each
# pattern in $SeriesLookup. Returns "Unknown" if no pattern matches.

function Get-MinifigSeries {
    param ([string]$MinifigID)

    foreach ($pattern in $SeriesLookup.Keys) {
        if ($MinifigID -match $pattern) {
            return $SeriesLookup[$pattern]
        }
    }

    return "Unknown"
}

# Filter the inventory to only items with ItemTypeID "M" (minifigures)
$minifigs = $bsx.BrickStoreXML.Inventory.Item |
    Where-Object { $_.ItemTypeID -eq "M" }

# Build a collection of result objects, one per minifig whose series is known
$results = foreach ($minifig in $minifigs) {
    $id     = $minifig.ItemID
    $series = Get-MinifigSeries $id

    # Skip minifigs that don't match any known series prefix
    if ($series -ne "Unknown") {
        [PSCustomObject]@{
            MinifigID = $id
            # Expand the single-character condition code to a readable label
            Condition = if ($minifig.Condition -eq 'U') { "Used" } else { "New" }
            Quantity  = $minifig.Qty
            Price     = $minifig.Price
            # Use an empty string if no remark is present; trim whitespace otherwise
            Remark    = if ($minifig.Remarks) { $minifig.Remarks.Trim() } else { "" }
            Series    = $series
        }
    }
}

# Sort results by Series then MinifigID, format as an auto-sized table, and
# write to the output file
$results | Sort-Object Series, MinifigID | Format-Table -AutoSize | Out-File -FilePath "$OutputDocx"
