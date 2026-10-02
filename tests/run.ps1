# Full check: classification test, then build an XLSX with every native chart and
# open it in Excel (no repair allowed). Chart pictures land in tests/out/png.
# Usage (from the project folder):  powershell -File tests\run.ps1
$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
$out = Join-Path $here "out"
New-Item -ItemType Directory -Force $out | Out-Null

node (Join-Path $here "harness.js") xlsx (Join-Path $out "charts.xlsx")
if ($LASTEXITCODE -ne 0) { Write-Host "Classification check failed." -ForegroundColor Red; exit 1 }

$result = & (Join-Path $here "excel_check.ps1") -File (Join-Path $out "charts.xlsx") -OutDir (Join-Path $out "png")
$result | ForEach-Object { Write-Host $_ }
if (-not ($result -match "^opened OK")) { Write-Host "Excel could not open the generated workbook." -ForegroundColor Red; exit 1 }
$expected = (Get-Content (Join-Path $out "charts.xlsx.names.txt")).Count
if (-not ($result -match "^opened OK: $expected chart objects")) { Write-Host "Excel loaded a different number of charts than were written ($expected)." -ForegroundColor Red; exit 1 }
Write-Host "All checks passed. Chart pictures: $out\png" -ForegroundColor Green
