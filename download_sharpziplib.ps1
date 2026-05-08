# Download ICSharpCode.SharpZipLib for NPOI (required for .xlsx write)
# Version: 1.3.3.11 (compatible with NPOI 2.5.6)
$outDir = "Assets\Plugins\ExcelToJsonPlugin\Dependencies"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$url = "https://globalcdn.nuget.org/packages/sharpziplib.1.3.3.nupkg"
$tmpZip = "$outDir\sz.nupkg"

Write-Host "Downloading SharpZipLib 1.3.3..."
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
(New-Object System.Net.WebClient).DownloadFile($url, $tmpZip)

# Extract
$tmpZip2 = "$outDir\sz.zip"
Move-Item $tmpZip $tmpZip2 -Force
Expand-Archive $tmpZip2 "$outDir\temp" -Force
Remove-Item $tmpZip2 -Force

# Copy DLL
Copy-Item "$outDir\temp\lib\netstandard2.0\ICSharpCode.SharpZipLib.dll" $outDir -Force
Remove-Item "$outDir\temp" -Recurse -Force

Write-Host "ICSharpCode.SharpZipLib.dll copied to $outDir" -ForegroundColor Green
Write-Host "Done. Restart Unity and try again."
