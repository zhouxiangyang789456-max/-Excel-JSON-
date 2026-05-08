# Download NPOI DLLs for Unity
# NPOI 2.7.0 - Pure C#, no System.Drawing dependency, Apache 2.0 License
# Output: Assets/Plugins/ExcelToJsonPlugin/Dependencies/

$ErrorActionPreference = "Stop"
$outDir = "Assets\Plugins\ExcelToJsonPlugin\Dependencies"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$versions = @(
    @{Version="2.7.0"; Url="https://www.nuget.org/api/v2/package/NPOI/2.7.0"}
    @{Version="2.7.0"; Url="https://globalcdn.nuget.org/packages/npoi.2.7.0.nupkg"}
)

$downloaded = $false
foreach ($v in $versions) {
    try {
        Write-Host "Trying: $($v.Url)"
        $tmpZip = "$outDir\npoi.nupkg"

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($v.Url, $tmpZip)

        # Extract DLLs from NuGet package (it's a ZIP, rename to .zip first)
        $tmpZip2 = "$outDir\npoi.zip"
        Move-Item -Path $tmpZip -Destination $tmpZip2 -Force
        Expand-Archive -Path $tmpZip2 -DestinationPath "$outDir\temp" -Force
        Remove-Item $tmpZip2 -Force

        # Copy relevant DLLs
        $libDir = "$outDir\temp\lib"
        if (Test-Path "$libDir\netstandard2.0") {
            Copy-Item "$libDir\netstandard2.0\*.dll" $outDir -Force
        } elseif (Test-Path "$libDir\netstandard2.1") {
            Copy-Item "$libDir\netstandard2.1\*.dll" $outDir -Force
        } elseif (Test-Path "$libDir\net40") {
            Copy-Item "$libDir\net40\*.dll" $outDir -Force
        }

        # Cleanup
        Remove-Item $tmpZip -Force
        Remove-Item "$outDir\temp" -Recurse -Force

        Write-Host "NPOI DLLs downloaded to $outDir" -ForegroundColor Green
        Get-ChildItem "$outDir\*.dll" | ForEach-Object { Write-Host "  $($_.Name)" }
        $downloaded = $true
        break
    } catch {
        Write-Host "Failed: $_" -ForegroundColor Yellow
    }
}

if (-not $downloaded) {
    Write-Host "`nAutomatic download failed. Please manually:" -ForegroundColor Red
    Write-Host "1. Visit https://www.nuget.org/packages/NPOI/2.7.0"
    Write-Host "2. Click 'Download package'"
    Write-Host "3. Rename .nupkg to .zip and extract"
    Write-Host "4. Copy netstandard2.0/*.dll to $outDir"
    Write-Host "`nRequired DLLs: NPOI.dll, NPOI.OOXML.dll, NPOI.OpenXml4Net.dll"
}
