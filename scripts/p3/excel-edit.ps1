# P3: edits a workbook through real Excel (COM) and saves it in place, in an
# Excel process this script starts and stops itself (never a user's instance).
# Usage: powershell -File excel-edit.ps1 -Path <xlsx> -EditsFile <json file, UTF-8>
#   json: [{ "cell": "B2", "value": "text or number" }, { "range": "A1:C1", "fill": "#C6EFCE" }]
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$EditsFile)
$ErrorActionPreference = 'Stop'
$list = [IO.File]::ReadAllText($EditsFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
Start-Sleep -Milliseconds 500
$own = @(Get-Process EXCEL | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($own.Count -ne 1) { throw "Excel did not start exactly one new process; refusing to touch an existing instance" }
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 3
  $wb = $xl.Workbooks.Open((Resolve-Path $Path).Path)
  $ws = $wb.Worksheets.Item(1)
  foreach ($e in $list) {
    if ($e.cell) {
      $v = $e.value
      # ConvertFrom-Json yields Decimal/Int64, which COM cannot pass to Excel
      if ($v -is [decimal] -or $v -is [int64] -or $v -is [int32]) { $v = [double]$v }
      # InvokeMember: PowerShell's COM binder caches the first value type per call site,
      # so assigning a number after a string through `.Value2 =` throws InvalidCast
      $cell = $ws.Range($e.cell)
      [void]$cell.GetType().InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $cell, @($v))
    }
    if ($e.range -and $e.fill) {
      $hex = $e.fill.TrimStart('#')
      $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
      $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
      $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
      $ws.Range($e.range).Interior.Color = $r + 256 * $g + 65536 * $b
    }
  }
  $wb.Save()
  $wb.Close($false)
  'saved'
} finally {
  try { $xl.Quit() } catch {}
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
  Start-Sleep -Seconds 2
  Stop-Process -Id $own[0] -Force -ErrorAction SilentlyContinue
}
