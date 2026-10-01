param([switch]$NoExcel)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
try {
 if (Test-Path -LiteralPath (Join-Path $root 'runtime/ollama/ollama.exe')) { try { & (Join-Path $PSScriptRoot 'Start-LocalAI.ps1') } catch { Write-Warning $_.Exception.Message } }
 $runtime=Get-Content -LiteralPath (Join-Path $root 'config/runtime.json') -Raw | ConvertFrom-Json
 if (!(Test-Path -LiteralPath $runtime.python)) { throw 'Không tìm thấy Python. Chạy Cai-dat-CCM.cmd trước.' }
 $cert=New-Object System.Security.Cryptography.X509Certificates.X509Certificate2((Join-Path $root 'config/localhost.cer'))
 if ($cert.NotAfter -lt (Get-Date).AddDays(7)) { throw 'Chứng chỉ HTTPS hết hạn/sắp hết hạn; cần cài lại chứng chỉ CCM.' }
 $trusted=Get-ChildItem Cert:\CurrentUser\Root | Where-Object Thumbprint -eq $cert.Thumbprint
 if (!$trusted) { throw 'Chứng chỉ CCM chưa được tin cậy. Chạy Cai-dat-CCM.cmd.' }
 function Health { try { return Invoke-RestMethod -Uri 'https://localhost:3443/api/health' -TimeoutSec 3 } catch { return $null } }
 $health=Health
 if ($health -and $health.installPath -ne $root) { throw ('Cổng 3443 đang chạy CCM từ thư mục khác: '+$health.installPath+'. Dừng bản đó trước khi dùng bản này.') }
 if (!$health) {
  $port=Get-NetTCPConnection -LocalPort 3443 -State Listen -ErrorAction SilentlyContinue
  if ($port) { throw 'Cổng 3443 bị dịch vụ khác chiếm hoặc HTTPS lỗi. Không dừng tiến trình khác tự động.' }
  $arg='"'+(Join-Path $root 'backend/run.py')+'"'
  $p=Start-Process -FilePath $runtime.python -ArgumentList $arg -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $root 'data/server-out.log') -RedirectStandardError (Join-Path $root 'data/server-error.log')
  $p.Id | Set-Content -LiteralPath (Join-Path $root 'data/server.pid')
  for ($i=0; $i -lt 30; $i++) { Start-Sleep -Milliseconds 500; $health=Health; if ($health) { break } }
 }
 if (!$health -or $health.app -ne 'CCM-AI') { throw 'Dịch vụ chưa sẵn sàng. Xem data/server-error.log.' }
 if (!$NoExcel) { Start-Process -FilePath 'excel.exe' }
} catch {
 Add-Type -AssemblyName PresentationFramework
 [System.Windows.MessageBox]::Show($_.Exception.Message,'CCM-AI – Không khởi động được') | Out-Null
 exit 1
}
