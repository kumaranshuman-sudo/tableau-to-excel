param([string]$File, [string]$OutDir)
# Opens an XLSX in Excel (no repair), reports chart count and exports every chart as PNG.
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force $OutDir | Out-Null
Get-ChildItem $OutDir -Filter *.png | Remove-Item -Force
$names = @()
if (Test-Path "$File.names.txt") { $names = Get-Content "$File.names.txt" }
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$xl.Visible = $false
try {
  # CorruptLoad = 0 (normal): a damaged package throws instead of silently repairing
  $wb = $xl.Workbooks.Open((Get-Item $File).FullName, 0, $true)
  $ws = $wb.Worksheets.Item(1)
  $n = $ws.ChartObjects().Count
  $shapes = $ws.Shapes.Count
  "opened OK: $n chart objects, $shapes shapes on '$($ws.Name)'"
  for ($i = 1; $i -le $shapes; $i++) {
    $sh = $ws.Shapes.Item($i)
    $label = if ($names.Count -ge $i) { $names[$i - 1] } else { $sh.Name }
    $safe = ($label -replace '[^A-Za-z0-9]+', '_')
    $kind = ""
    try { $kind = $sh.Chart.ChartType } catch { $kind = "type n/a" }
    $png = Join-Path $OutDir ("{0:D2}_{1}.png" -f $i, $safe)
    $sh.Top = 0; $sh.Left = 0                      # off-screen charts export as empty files
    for ($try = 0; $try -lt 4; $try++) {
      try { $null = $sh.Chart.Export($png) } catch { "  export failed for ${label}: $($_.Exception.Message)"; break }
      if ((Test-Path $png) -and (Get-Item $png).Length -gt 0) { break }
      Start-Sleep -Milliseconds 400
    }
    "  [$i] $label  (ChartType=$kind, HasChart=$($sh.HasChart))"
  }
  $wb.Close($false)
} catch {
  "OPEN FAILED: $($_.Exception.Message)"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

