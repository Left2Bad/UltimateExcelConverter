param(
    [string]$Python = '.venv\Scripts\python.exe'
)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
& $Python -m PyInstaller --noconfirm --onedir --windowed --name ExcelConverter --add-data 'excel_converter/profiles;excel_converter/profiles' --exclude-module numpy --exclude-module pandas --exclude-module PIL --exclude-module lxml desktop.py
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
Copy-Item -LiteralPath README.md -Destination dist\ExcelConverter\README.md
Write-Host 'Ready: dist\ExcelConverter\ExcelConverter.exe. Copy the entire ExcelConverter directory.'
