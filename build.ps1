[CmdletBinding()]
param(
    [string]$OutputName = "accessMenu.nvda-addon"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$rootDir = Split-Path -Parent $PSCommandPath
$addonDir = Join-Path $rootDir "addon"
if ([System.IO.Path]::IsPathRooted($OutputName)) {
    $outputPath = $OutputName
} else {
    $outputPath = Join-Path $rootDir $OutputName
}
$zipPath = $outputPath

# Compress-Archive only supports .zip output; build the zip then rename if needed.
if ([System.IO.Path]::GetExtension($zipPath).ToLowerInvariant() -ne ".zip") {
    $zipPath = "$outputPath.zip"
}

if (-not (Test-Path -LiteralPath $addonDir)) {
    throw "Expected addon directory not found: $addonDir"
}

if (Test-Path -LiteralPath $outputPath) {
    Remove-Item -LiteralPath $outputPath -Force
}
if ($zipPath -ne $outputPath -and (Test-Path -LiteralPath $zipPath)) {
    Remove-Item -LiteralPath $zipPath -Force
}

Push-Location -LiteralPath $addonDir
try {
    # Zip the *contents* of addon/ so manifest.ini sits at archive root.
    Compress-Archive -Path * -DestinationPath $zipPath -CompressionLevel Optimal
} finally {
    Pop-Location
}

if ($zipPath -ne $outputPath) {
    Move-Item -LiteralPath $zipPath -Destination $outputPath -Force
}

$builtName = Split-Path -Leaf $outputPath
$builtDir = Split-Path -Parent $outputPath
Write-Host "Built $builtName in $builtDir"
