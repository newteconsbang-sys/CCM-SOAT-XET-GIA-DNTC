param([string]$PythonPath,[switch]$NoShortcut,[switch]$SkipTrust)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
try {
 if (!$PythonPath) {
  $bundled=Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
  if (Test-Path -LiteralPath (Join-Path $root 'runtime/python/python.exe')) { $PythonPath=Join-Path $root 'runtime/python/python.exe' }
  elseif (Test-Path -LiteralPath $bundled) { $PythonPath=$bundled }
  else { $PythonPath=(Get-Command python.exe -ErrorAction Stop).Source }
 }
 if (!(Test-Path -LiteralPath (Join-Path $root 'frontend/dist/index.html'))) { throw 'Thiếu bản frontend đã build. Dùng bộ bàn giao đầy đủ.' }
 & $PythonPath -c "import sys;sys.path.insert(0,r'$root\backend\vendor');import fastapi,uvicorn,cryptography,openpyxl,docx,reportlab,pypdf"
 if ($LASTEXITCODE -ne 0) { throw 'Python thiếu thư viện. Dùng bộ runtime hoặc cài backend/requirements.txt.' }
 & $PythonPath (Join-Path $root 'backend/certificate.py')
 if ($LASTEXITCODE -ne 0) { throw 'Không tạo được chứng chỉ.' }
 $certFile=Join-Path $root 'config/localhost.cer'
 $cert=New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFile)
 @{python=$PythonPath;root=$root;port=3443;certificate=$cert.Thumbprint} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'config/runtime.json') -Encoding UTF8
 # Per-user sideload registration, as used by Microsoft's office-addin-dev-settings.
 $reg='HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
 New-Item -Path $reg -Force | Out-Null
 New-ItemProperty -Path $reg -Name '79e68a1b-4a3b-41d2-970b-f2d83c16dc01' -Value (Join-Path $root 'manifests/ccm-ai.xml') -PropertyType String -Force | Out-Null
 if (!$NoShortcut) {
  $shell=New-Object -ComObject WScript.Shell
  $desktop=[Environment]::GetFolderPath('Desktop')
  $asciiShortcut=Join-Path $desktop 'Khoi-dong-CCM-AI.lnk'
  $shortcut=$shell.CreateShortcut($asciiShortcut)
  $shortcut.TargetPath=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
  $shortcut.Arguments='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $root 'scripts/Start-CCM.ps1')+'"'
  $shortcut.WorkingDirectory=$root;$shortcut.WindowStyle=7;$shortcut.Description='CCM-AI local HTTPS and Excel';$shortcut.Save()
  $displayShortcut=Join-Path $desktop 'Khởi động CCM-AI.lnk'
  Move-Item -LiteralPath $asciiShortcut -Destination $displayShortcut -Force
 }
 if (!$SkipTrust -and !(Get-ChildItem Cert:\CurrentUser\Root | Where-Object Thumbprint -eq $cert.Thumbprint)) {
  Write-Host 'Windows sẽ hỏi tin cậy chứng chỉ CCM-AI localhost. Vui lòng xác nhận trực tiếp.'
  & certutil.exe -user -addstore Root $certFile
  if ($LASTEXITCODE -ne 0) { throw 'Chưa tin cậy chứng chỉ. Cấu hình và shortcut đã tạo; chạy lại bộ cài khi sẵn sàng.' }
 }
 Write-Host 'Đã cấu hình sideload cá nhân và shortcut. Lưu workbook rồi đóng/mở Excel.'
 Write-Host 'Trong Excel: Home > Add-ins > Advanced > My Add-ins > Developer Add-ins > CCM-AI > Add.'
 Write-Host 'Nếu doanh nghiệp chặn Developer Add-ins, xem docs/CAI-DAT.md về catalog và triển khai qua IT.'
} catch { Write-Host $_.Exception.Message -ForegroundColor Red; exit 1 }
