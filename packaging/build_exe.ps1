$ErrorActionPreference = "Stop"

$root = Resolve-Path "$PSScriptRoot\\.."
Set-Location $root

python -m pip install --upgrade pyinstaller

$env:PLAYWRIGHT_BROWSERS_PATH = "$root\\playwright-browsers"
python -m playwright install chromium

$addData = @()
$browsersPath = "$root\\playwright-browsers"
if (Test-Path $browsersPath) {
  $addData += @("--add-data", "$browsersPath;playwright-browsers")
}
$fontsPath = "$root\\assets\\fonts"
if (Test-Path $fontsPath) {
  $addData += @("--add-data", "$fontsPath;assets\\fonts")
}

pyinstaller `
  --noconfirm `
  --clean `
  --name "DIRGC-Automation" `
  --windowed `
  --onefile `
  --collect-all qfluentwidgets `
  --collect-all PyQt5 `
  --collect-all playwright `
  @addData `
  run_dirgc_gui.py
