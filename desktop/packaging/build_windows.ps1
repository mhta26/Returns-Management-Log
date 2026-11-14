$ErrorActionPreference = 'Stop'
Set-Location (Split-Path $PSScriptRoot -Parent)
python -m PyInstaller --noconfirm --clean --windowed --name ReturnsManager --add-data 'returns_app/assets;returns_app/assets' --collect-all PySide6.QtPdf --collect-all tzdata --collect-all arabic_reshaper --collect-all bidi launch_client.py
if ($LASTEXITCODE -ne 0) { throw 'Client build failed' }
python -m PyInstaller --noconfirm --clean --console --name ReturnsServer --add-data 'returns_app/assets;returns_app/assets' --collect-all arabic_reshaper --collect-all bidi --collect-all tzdata launch_server.py
if ($LASTEXITCODE -ne 0) { throw 'Server build failed' }
& '.\dist\ReturnsServer\ReturnsServer.exe' --self-test
if ($LASTEXITCODE -ne 0) { throw 'Packaged server smoke test failed' }
$clientTest = Start-Process -FilePath '.\dist\ReturnsManager\ReturnsManager.exe' -ArgumentList '--self-test' -PassThru -Wait
if ($clientTest.ExitCode -ne 0) { throw 'Packaged client smoke test failed' }
$compiler = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
if (Test-Path $compiler) {
  & $compiler 'packaging\installer.iss'
  if ($LASTEXITCODE -ne 0) { throw 'Installer build failed' }
}
New-Item -ItemType Directory -Force release | Out-Null
Compress-Archive -Path 'dist\ReturnsManager','dist\ReturnsServer','README.md','OPERATIONS.md','THIRD_PARTY.md' -DestinationPath 'release\ReturnsManager-Portable-Windows-x64.zip' -Force
Get-ChildItem release -File | Get-FileHash -Algorithm SHA256 | ForEach-Object { "$($_.Hash)  $(Split-Path $_.Path -Leaf)" } | Set-Content 'release\SHA256SUMS.txt'
