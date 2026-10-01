$repo = $PSScriptRoot
$env:RUSTUP_HOME = Join-Path $repo '.local-tools\rustup'
$env:CARGO_HOME = Join-Path $repo '.local-tools\cargo'
$env:PATH = "$repo\.local-tools\node-v24.14.0-win-x64;$repo\.local-tools\cargo\bin;$repo\.local-tools\w64devkit\bin;$env:PATH"
$env:CARGO_BUILD_TARGET = 'x86_64-pc-windows-gnu'
$env:LIBRARY_PATH = "$repo\.local-tools\rustup\toolchains\stable-x86_64-pc-windows-gnu\lib\rustlib\x86_64-pc-windows-gnu\lib\self-contained"
if (-not $env:BUILD_DIR) { $env:BUILD_DIR = Join-Path $repo 'apps\shell\release\dvh-home' }
Set-Location $repo
& npm.cmd run dist:win
exit $LASTEXITCODE


