$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
& (Join-Path $PSScriptRoot 'Stop-CCM.ps1')
Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer' -Name '79e68a1b-4a3b-41d2-970b-f2d83c16dc01' -ErrorAction SilentlyContinue
$shortcut=Join-Path ([Environment]::GetFolderPath('Desktop')) 'Khởi động CCM-AI.lnk'
if (Test-Path -LiteralPath $shortcut) { Remove-Item -LiteralPath $shortcut }
$config=Join-Path $root 'config/runtime.json'
if (Test-Path -LiteralPath $config) {
 $ccmConfig=Get-Content -LiteralPath $config -Raw | ConvertFrom-Json
 $cert=Get-ChildItem Cert:\CurrentUser\Root | Where-Object Thumbprint -eq $ccmConfig.certificate
 if ($cert) { Remove-Item -LiteralPath $cert.PSPath }
}
Write-Host 'Đã gỡ đăng ký và shortcut; giữ nguyên hồ sơ, thư viện, báo cáo, bản sao.'
