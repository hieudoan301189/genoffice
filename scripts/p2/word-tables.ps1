# P2: lists the content controls and tables of a docx as real Word sees them
# (read-only open, never saved) in a Word process this script starts itself.
# Usage: powershell -File word-tables.ps1 -Path <docx> -Out <json>
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Out)
$ErrorActionPreference = 'Stop'
$wordExe = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe').'(default)'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
Start-Process -FilePath $wordExe -ArgumentList '/a', '/automation', '-Embedding'
Start-Sleep -Seconds 4
$word = New-Object -ComObject Word.Application
$own = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($own.Count -ne 1) { throw "Word did not start exactly one new process; refusing to touch an existing instance" }
try {
  $word.Visible = $false; $word.DisplayAlerts = 0
  $doc = $word.Documents.Open((Resolve-Path $Path).Path, $false, $true, $false)
  $controls = @()
  foreach ($cc in $doc.ContentControls) {
    $controls += [ordered]@{ tag = $cc.Tag; type = $cc.Type; tables = $cc.Range.Tables.Count; mapped = $cc.XMLMapping.IsMapped }
  }
  $tables = @()
  foreach ($tb in $doc.Tables) {
    $rows = @()
    foreach ($row in $tb.Rows) {
      $cells = @()
      foreach ($cell in $row.Cells) { $cells += ($cell.Range.Text -replace '[\r\a]', '').Trim() }
      $rows += , $cells
    }
    $tables += [ordered]@{
      rows = $tb.Rows.Count
      headingRow = [bool]$tb.Rows.Item(1).HeadingFormat
      firstFill = $tb.Cell(1, 1).Shading.BackgroundPatternColor
      text = $rows
    }
  }
  $doc.Close(0)
  $json = [ordered]@{ controls = $controls; tables = $tables } | ConvertTo-Json -Depth 6
  [IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding $false))
  $json
} finally {
  try { $word.Quit(0) } catch {}
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
  Start-Sleep -Seconds 2
  Stop-Process -Id $own[0] -Force -ErrorAction SilentlyContinue
}
