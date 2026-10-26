$ErrorActionPreference = "Stop"

python -m pip install --upgrade pyinstaller
if (Test-Path -LiteralPath "build") { Remove-Item -LiteralPath "build" -Recurse -Force }
if (Test-Path -LiteralPath "dist") { Remove-Item -LiteralPath "dist" -Recurse -Force }

python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --noconsole `
  --name GooddeggInventory `
  --add-data "assets;assets" `
  desktop_app.py

Write-Host "Built dist\GooddeggInventory.exe"
