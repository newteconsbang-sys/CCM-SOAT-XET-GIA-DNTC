param([switch]$NoBrowser)
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'Start-CCM.ps1') -NoExcel -NoAI -Headless
try {
    $json = curl.exe -s -k https://localhost:3443/api/health
    if (!$json) { throw 'Dich vu CCM-AI chua san sang.' }
    $health = $json | ConvertFrom-Json
    if ($health.app -ne 'CCM-AI') { throw 'Dich vu CCM-AI chua san sang.' }
    if (!$NoBrowser) { Start-Process 'https://localhost:3443/' }
}
catch { Write-Error $_.Exception.Message; exit 1 }
