# P3 demo helper, started elevated once (Bandicam requires administrator).
# Bandicam 4.1 has no working command-line switches, so recording is toggled
# with its record hotkey (F12 by default, dwVideoHotkey), sent from this
# elevated process so the elevated Bandicam receives it.
#
# 1. Finds the Bandicam recording mode that records the screen: the user's
#    settings may have no mode selected (nTargetMode = -1), and /record then does
#    nothing. Each candidate mode records 4 s; the first that writes a video wins
#    (the test video is moved into -SignalDir, never deleted).
# 2. Starts Bandicam in that mode and runs /record, /stop when the demo drops
#    record.signal / stop.signal into -SignalDir; shutdown.signal closes it.
# 3. After Bandicam has exited (it saves its settings on exit) the user's own
#    values are written back: mode, sub-mode and recording rectangle.
# Writes ready.signal / *.done / shutdown.done; gives up after 15 minutes.
param(
  [Parameter(Mandatory = $true)][string]$SignalDir,
  [Parameter(Mandatory = $true)][string]$RectBackup,
  [string]$Bandicam = 'C:\Program Files (x86)\Bandicam\bdcam.exe'
)
$ErrorActionPreference = 'Stop'
$key = 'HKCU:\Software\BANDISOFT\BANDICAM\OPTION'
$deadline = (Get-Date).AddMinutes(15)
$opt = Get-ItemProperty $key
$outDir = $opt.sOutputFolder
$original = @{
  nTargetMode = [uint32]$opt.nTargetMode
  nScreenRecordingSubMode = [uint32]$opt.nScreenRecordingSubMode
  bStartMinimized = [uint32]$opt.bStartMinimized
}
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class Keys { [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra); }
'@
$hotkey = [byte]($opt.dwVideoHotkey -band 0xFF)
if (-not $hotkey) { $hotkey = [byte]0x7B }
function Toggle-Recording {
  [Keys]::keybd_event($hotkey, 0, 0, [UIntPtr]::Zero)
  Start-Sleep -Milliseconds 80
  [Keys]::keybd_event($hotkey, 0, 2, [UIntPtr]::Zero)
}
function Log($text) { Add-Content -Path (Join-Path $SignalDir 'helper.log') -Value "$(Get-Date -Format o) $text" }

function Close-Bandicam {
  foreach ($p in @(Get-Process bdcam -ErrorAction SilentlyContinue)) { [void]$p.CloseMainWindow() }
  for ($i = 0; $i -lt 20 -and (Get-Process bdcam -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }
  Get-Process bdcam -ErrorAction SilentlyContinue | Stop-Process -Force
  Start-Sleep -Seconds 1
}

function Restore-Settings {
  foreach ($name in $original.Keys) { Set-ItemProperty $key -Name $name -Value $original[$name] -Type DWord }
  $rect = (Get-Content $RectBackup -Raw).Trim().Split(',')
  $names = 'TargetRect.left', 'TargetRect.top', 'TargetRect.right', 'TargetRect.bottom'
  for ($i = 0; $i -lt 4; $i++) { Set-ItemProperty $key -Name $names[$i] -Value ([uint32]$rect[$i]) -Type DWord }
  Log 'settings restored'
}

function Newest { Get-ChildItem $outDir -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }

function Start-Bandicam($mode) {
  Set-ItemProperty $key -Name nTargetMode -Value ([uint32]$mode) -Type DWord
  # a minimized main window: the hotkey still records, nothing covers the demo
  Set-ItemProperty $key -Name bStartMinimized -Value 1 -Type DWord
  Set-ItemProperty $key -Name nScreenRecordingSubMode -Value 0 -Type DWord
  Start-Process -FilePath $Bandicam -ArgumentList '/nosplash'
  Start-Sleep -Seconds 6
}

try {
  if (Get-Process bdcam -ErrorAction SilentlyContinue) { Close-Bandicam }
  # 1. which mode records the screen (rectangle = whole screen, set by the demo)
  $found = $null
  foreach ($mode in @($original.nTargetMode, 0, 1, 2)) {
    $before = Newest
    Start-Bandicam $mode
    Toggle-Recording
    Start-Sleep -Seconds 4
    Toggle-Recording
    Start-Sleep -Seconds 4
    Close-Bandicam
    $after = Newest
    if ($after -and (-not $before -or $after.FullName -ne $before.FullName)) {
      Move-Item $after.FullName (Join-Path $SignalDir "calibration-mode-$mode$($after.Extension)")
      $found = $mode
      Log "mode $mode records"
      break
    }
    Log "mode $mode wrote nothing"
  }
  if ($null -eq $found) {
    Set-Content -Path (Join-Path $SignalDir 'failed.signal') -Value 'no recording mode wrote a video'
    return
  }

  # 2. the demo recording
  Start-Bandicam $found
  Set-Content -Path (Join-Path $SignalDir 'ready.signal') -Value "$found"
  $handled = @{}
  while ((Get-Date) -lt $deadline) {
    foreach ($name in 'record', 'stop') {
      $file = Join-Path $SignalDir "$name.signal"
      if ((Test-Path $file) -and -not $handled[$name]) {
        $handled[$name] = $true
        Toggle-Recording
        Set-Content -Path (Join-Path $SignalDir "$name.done") -Value 'done'
        Log "$name sent"
      }
    }
    if (Test-Path (Join-Path $SignalDir 'shutdown.signal')) { break }
    Start-Sleep -Milliseconds 300
  }
  if ($handled['record'] -and -not $handled['stop']) { Toggle-Recording; Start-Sleep -Seconds 5 }
} finally {
  # 3. Bandicam saves its settings on exit: close it, then write the user's values back
  if (Get-Process bdcam -ErrorAction SilentlyContinue) { Close-Bandicam }
  Restore-Settings
  Set-Content -Path (Join-Path $SignalDir 'shutdown.done') -Value 'done'
}
