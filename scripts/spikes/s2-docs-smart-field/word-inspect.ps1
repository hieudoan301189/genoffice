# Spike S2: lists the content controls of a docx as real Word sees them (read-only
# open, never saved) in a Word process this script starts without add-ins.
# Usage: powershell -File word-inspect.ps1 -Path <docx> -Out <json>
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
    $controls += [ordered]@{ tag = $cc.Tag; text = $cc.Range.Text; mapped = $cc.XMLMapping.IsMapped; xpath = $cc.XMLMapping.XPath }
  }
  $paragraphs = @()
  foreach ($p in $doc.Paragraphs) { $t = $p.Range.Text.Trim(); if ($t -and $t.Length -lt 120) { $paragraphs += $t } }
  $doc.Close(0)
  $json = [ordered]@{ controls = $controls; paragraphs = $paragraphs } | ConvertTo-Json -Depth 5
  [IO.File]::WriteAllText($Out, $json, (New-Object Text.UTF8Encoding $false))
  $json
} finally {
  try { $word.Quit(0) } catch {}
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
  Start-Sleep -Seconds 2
  Stop-Process -Id $own[0] -Force -ErrorAction SilentlyContinue
}
