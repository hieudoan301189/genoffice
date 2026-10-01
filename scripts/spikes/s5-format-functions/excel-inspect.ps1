# Spike S5: reads cell formats of a saved workbook through real Excel (COM), in a
# process this script starts and stops itself.
# Usage: powershell -File excel-inspect.ps1 -Path <xlsx> -Cells F1,G2,... -Out <json>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [Parameter(Mandatory = $true)][string[]]$Cells,
  [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'
# -File passes "F1,G2" as one string
$Cells = @($Cells | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
function Hex($ole) { if ($null -eq $ole) { return $null }; $v = [int64]$ole; '#{0:X2}{1:X2}{2:X2}' -f ($v -band 255), (($v -shr 8) -band 255), (($v -shr 16) -band 255) }

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
Start-Sleep -Milliseconds 500
$own = @(Get-Process EXCEL | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($own.Count -ne 1) { throw "Excel did not start exactly one new process; refusing to touch an existing instance" }
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 3
  try { $wb = $xl.Workbooks.Open((Resolve-Path $Path).Path, 0, $true) } catch { throw "open failed: $($_.Exception.Message)" }
  $ws = $wb.Worksheets.Item(1)
  $result = [ordered]@{}
  foreach ($a in $Cells) {
    $r = $ws.Range($a)
    $spill = $null; try { $spill = $r.HasSpill } catch {}
    $f2 = $null; try { $f2 = $r.Formula2 } catch { $f2 = $r.Formula }
    $result[$a] = [ordered]@{
      text = $r.Text
      bold = $r.Font.Bold
      italic = $r.Font.Italic
      underline = $r.Font.Underline
      fontColor = Hex $r.Font.Color
      fill = $(if ($r.Interior.ColorIndex -eq -4142) { $null } else { Hex $r.Interior.Color })
      numberFormat = $r.NumberFormat
      formula = $f2
      hasSpill = $spill
    }
  }
  $wb.Close($false)
  [IO.File]::WriteAllText($Out, ($result | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding $false))
  Get-Content -Raw $Out
} finally {
  try { $xl.Quit() } catch {}
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
  Start-Sleep -Seconds 2
  Stop-Process -Id $own[0] -Force -ErrorAction SilentlyContinue
}
