# Spike S1: open, inspect, edit and save the probe documents with real Word/Excel
# through COM. Only Office processes started by this script are touched; an
# instance the user already has open is never reused or closed.
# Usage: powershell -File office-roundtrip.ps1 -OutDir <dir> -Only word|excel
param(
  [Parameter(Mandatory = $true)][string]$OutDir,
  [Parameter(Mandatory = $true)][ValidateSet('word', 'excel')][string]$Only
)
$ErrorActionPreference = 'Stop'
$OutDir = (Resolve-Path $OutDir).Path
$config = Get-Content -Raw -Encoding UTF8 (Join-Path $OutDir 'office-config.json') | ConvertFrom-Json
$report = [ordered]@{ word = @(); excel = @() }
$logPath = Join-Path $OutDir "office-progress-$Only.log"
Set-Content -Path $logPath -Value '' -Encoding ASCII
function Log([string]$msg) { Add-Content -Path $logPath -Value ((Get-Date -Format 'HH:mm:ss.fff') + ' ' + $msg) -Encoding ASCII }

# $noAddinsExe: pre-launch that executable with /a (no add-ins, no Normal.dotm, no
# settings writes) as a COM server, so the next CoCreateInstance binds to it. COM
# add-ins installed on this machine (DVH-Word, Office Tab...) otherwise load into
# automation instances too, and one of them blocks SaveAs2.
function Start-OwnInstance([string]$progId, [string]$processName, [string]$noAddinsExe = '') {
  $before = @(Get-Process $processName -ErrorAction SilentlyContinue | ForEach-Object Id)
  if ($noAddinsExe) {
    Start-Process -FilePath $noAddinsExe -ArgumentList '/a', '/automation', '-Embedding'
    Start-Sleep -Seconds 5
  }
  $app = New-Object -ComObject $progId
  Start-Sleep -Milliseconds 500
  $after = @(Get-Process $processName -ErrorAction SilentlyContinue | ForEach-Object Id)
  $own = @($after | Where-Object { $before -notcontains $_ })
  if ($own.Count -ne 1) {
    throw "$progId did not start exactly one new process (got $($own.Count)); refusing to touch an existing instance"
  }
  return @{ App = $app; Pid = $own[0] }
}

function Stop-OwnInstance($instance, [switch]$Word) {
  # Word: Quit(0) = wdDoNotSaveChanges, so a dirty document never raises a save prompt
  try { if ($Word) { $instance.App.Quit(0) } else { $instance.App.Quit() } } catch {}
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($instance.App)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  for ($i = 0; $i -lt 20; $i++) {
    if (-not (Get-Process -Id $instance.Pid -ErrorAction SilentlyContinue)) { return }
    Start-Sleep -Milliseconds 250
  }
  Stop-Process -Id $instance.Pid -Force -ErrorAction SilentlyContinue
}

function Get-Parts($parts) {
  $list = @()
  foreach ($p in $parts) {
    if ($p.BuiltIn) { continue }
    $list += [ordered]@{ ns = $p.NamespaceURI; id = $p.Id; xmlChars = $p.XML.Length }
  }
  return $list
}

# ---------------- Word ----------------
if ($Only -eq 'word') {
$wordExe = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe').'(default)'
$word = Start-OwnInstance 'Word.Application' 'WINWORD' $wordExe
Log "word started pid $($word.Pid) without add-ins"
$watchdog = Start-Job -ArgumentList $word.Pid, $logPath -ScriptBlock {
  param($ownPid, $log)
  Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
  $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $ownPid)
  $seen = @{}
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Seconds 2
    if (-not (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { return }
    foreach ($w in [System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)) {
      $texts = @($w.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name } | Where-Object { $_ })
      $line = 'DIALOG ' + $w.Current.Name + ' :: ' + ($texts -join ' | ')
      if (-not $seen[$line]) { $seen[$line] = $true; Add-Content -Path $log -Value ((Get-Date -Format 'HH:mm:ss.fff') + ' ' + $line) -Encoding ASCII }
    }
  }
  Add-Content -Path $log -Value ((Get-Date -Format 'HH:mm:ss.fff') + ' WATCHDOG timeout: killing own Word pid ' + $ownPid) -Encoding ASCII
  Stop-Process -Id $ownPid -Force -ErrorAction SilentlyContinue
}
try {
  $word.App.Visible = $false
  $word.App.DisplayAlerts = 0
  foreach ($name in @('dvh-s1.docx', 'dvh-s1-big.docx')) {
    $src = Join-Path $OutDir $name
    $dst = Join-Path $OutDir ("word-saved-" + $name)
    $entry = [ordered]@{ file = $name }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Log "word open $name"
    $doc = $word.App.Documents.Open($src, $false, $false, $false)
    $entry.openMs = $sw.ElapsedMilliseconds
    Log "word opened $name in $($entry.openMs) ms"
    $controls = @()
    foreach ($cc in $doc.ContentControls) {
      $mapping = $cc.XMLMapping
      $controls += [ordered]@{
        tag = $cc.Tag; title = $cc.Title; type = $cc.Type; text = $cc.Range.Text
        mapped = $mapping.IsMapped; xpath = $mapping.XPath
      }
    }
    $entry.controlsOnOpen = $controls
    $entry.partsOnOpen = Get-Parts $doc.CustomXMLParts
    Log 'word read controls/parts'
    if ($name -eq 'dvh-s1.docx') {
      Log 'word edit content control'
      foreach ($cc in $doc.ContentControls) {
        if ($cc.Tag -eq $config.projectTag) { $cc.Range.Text = $config.wordEditValue }
      }
      $model = $doc.CustomXMLParts.SelectByNamespace($config.modelNs).Item(1)
      $model.NamespaceManager.AddNamespace('dvh', $config.modelNs)
      $entry.modelProjectAfterEdit = $model.SelectSingleNode($config.projectXPath).Text
      $doc.Content.InsertParagraphAfter()
      $doc.Content.InsertAfter('Paragraph added in Word.')
    }
    $sw.Restart()
    Log "word save $name"
    # SaveAs2 with positional args hangs from PowerShell on this machine (no dialog,
    # with or without add-ins); the legacy SaveAs with ByRef args saves in ~0.3 s.
    [object]$savePath = [string]$dst
    [object]$saveFormat = 16
    $doc.SaveAs([ref]$savePath, [ref]$saveFormat)
    $entry.saveMs = $sw.ElapsedMilliseconds
    Log "word saved in $($entry.saveMs) ms"
    $doc.Close(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
    $report.word += $entry
  }
} finally {
  Stop-OwnInstance $word -Word
  Stop-Job $watchdog -ErrorAction SilentlyContinue; Remove-Job $watchdog -Force -ErrorAction SilentlyContinue
}
}

# ---------------- Excel ----------------
if ($Only -eq 'excel') {
$excel = Start-OwnInstance 'Excel.Application' 'EXCEL'
try {
  $excel.App.Visible = $false
  $excel.App.DisplayAlerts = $false
  $excel.App.AutomationSecurity = 3
  foreach ($name in @('dvh-s1.xlsx', 'dvh-s1-big.xlsx')) {
    $src = Join-Path $OutDir $name
    $dst = Join-Path $OutDir ("excel-saved-" + $name)
    $entry = [ordered]@{ file = $name }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $wb = $excel.App.Workbooks.Open($src, 0, $false)
    $entry.openMs = $sw.ElapsedMilliseconds
    $ws = $wb.Worksheets.Item(1)
    $names = @()
    foreach ($n in $wb.Names) { $names += [ordered]@{ name = $n.Name; refersTo = $n.RefersTo; visible = $n.Visible } }
    $entry.namesOnOpen = $names
    $excel.App.CalculateFull()
    $entry.c5ViaHiddenName = $ws.Range('C5').Text
    $entry.d5ViaVisibleName = $ws.Range('D5').Text
    $entry.partsOnOpen = Get-Parts $wb.CustomXMLParts
    if ($name -eq 'dvh-s1.xlsx') {
      [void]$ws.Rows.Item(3).Insert()
      $after = @()
      foreach ($n in $wb.Names) { $after += [ordered]@{ name = $n.Name; refersTo = $n.RefersTo; visible = $n.Visible } }
      $entry.namesAfterInsertRow3 = $after
    }
    $sw.Restart()
    $wb.SaveAs($dst, 51)
    $entry.saveMs = $sw.ElapsedMilliseconds
    $wb.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)
    $report.excel += $entry
  }
} finally {
  Stop-OwnInstance $excel
}
}

$json = $report[$Only] | ConvertTo-Json -Depth 8
[IO.File]::WriteAllText((Join-Path $OutDir "office-report-$Only.json"), $json, (New-Object Text.UTF8Encoding $false))
Write-Output $json
