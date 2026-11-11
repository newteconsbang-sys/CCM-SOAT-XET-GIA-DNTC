param([switch]$NoExcel,[switch]$NoAI,[switch]$Headless)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
try {
 $runtime=Get-Content -LiteralPath (Join-Path $root 'config/runtime.json') -Raw | ConvertFrom-Json
 if (!(Test-Path -LiteralPath $runtime.python)) { throw 'Khong tim thay Python. Chay Cai-dat-CCM.cmd truoc.' }
 $cert=New-Object System.Security.Cryptography.X509Certificates.X509Certificate2((Join-Path $root 'config/localhost.cer'))
 if ($cert.NotAfter -lt (Get-Date).AddDays(7)) { throw 'Chung chi HTTPS het han/sap het han; can cai lai chung chi CCM.' }
 $trusted=Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $cert.Thumbprint -or $_.Subject -like '*CCM-AI*' }
 if (!$trusted) { throw 'Chung chi CCM chua duoc tin cay. Chay Cai-dat-CCM.cmd.' }
 function Health {
  try {
   $json = curl.exe -s -k https://localhost:3443/api/health
   if ($json) { return ($json | ConvertFrom-Json) }
  } catch {}
  return $null
 }
 $health=Health
 if ($health -and $health.installPath -ne $root) { throw ('Cong 3443 dang chay CCM tu thu muc khac: '+$health.installPath+'. Dung ban do truoc khi dung ban nay.') }
 # Dich vu dang chay ban cu (sau khi cap nhat ma nguon) thi tu tat va bat lai ban moi.
 $versionFile=Join-Path $root 'backend/VERSION'
 if ($health -and (Test-Path -LiteralPath $versionFile)) {
  $want=(Get-Content -LiteralPath $versionFile -Raw).Trim()
  if ($health.version -ne $want) {
   & (Join-Path $PSScriptRoot 'Stop-CCM.ps1')
   $health=Health
   if ($health) { throw ('CCM ban cu ('+$health.version+') van dang chay, khong tat duoc. Khoi dong lai may roi mo lai CCM-AI.') }
  }
 }
 if (!$health) {
  $port=Get-NetTCPConnection -LocalPort 3443 -State Listen -ErrorAction SilentlyContinue
  if ($port) { throw 'Cong 3443 bi dich vu khac chiem hoac HTTPS loi. Khong dung tien trinh khac tu dong.' }
  $arg='"'+(Join-Path $root 'backend/run.py')+'"'
  $p=Start-Process -FilePath $runtime.python -ArgumentList $arg -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $root 'data/server-out.log') -RedirectStandardError (Join-Path $root 'data/server-error.log')
  $p.Id | Set-Content -LiteralPath (Join-Path $root 'data/server.pid')
  for ($i=0; $i -lt 30; $i++) { Start-Sleep -Milliseconds 500; $health=Health; if ($health) { break } }
 }
 if (!$health -or $health.app -ne 'CCM-AI') { throw 'Dich vu chua san sang. Xem data/server-error.log.' }
 if (!$NoExcel) { Start-Process -FilePath 'excel.exe' }
 if (!$NoAI -and (Test-Path -LiteralPath (Join-Path $root 'runtime/ollama/ollama.exe'))) { try { & (Join-Path $PSScriptRoot 'Start-LocalAI.ps1') } catch { Write-Warning $_.Exception.Message } }
} catch {
 if ($Headless) { $_.Exception.Message | Add-Content -LiteralPath (Join-Path $root 'data/startup-error.log'); exit 1 }
 Add-Type -AssemblyName PresentationFramework
 [System.Windows.MessageBox]::Show($_.Exception.Message,'CCM-AI – Khong khoi dong duoc') | Out-Null
 exit 1
}
