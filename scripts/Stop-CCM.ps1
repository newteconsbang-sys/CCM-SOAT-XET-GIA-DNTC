$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$runPy=($root+'\backend\run.py').ToLower()
function Test-CcmProcess($procId) {
 $process=Get-CimInstance Win32_Process -Filter "ProcessId=$procId" -ErrorAction SilentlyContinue
 return ($process -and $process.CommandLine -and $process.CommandLine.Replace('/','\').ToLower().Contains($runPy))
}
$ids=@()
$pidFile=Join-Path $root 'data/server.pid'
if (Test-Path -LiteralPath $pidFile) {
 $ccmPid=[int](Get-Content -LiteralPath $pidFile)
 if (Test-CcmProcess $ccmPid) { $ids+=$ccmPid }
}
# Du phong: tien trinh dang giu cong 3443 ma chinh la CCM cua thu muc nay (file pid bi cu).
$listen=Get-NetTCPConnection -LocalPort 3443 -State Listen -ErrorAction SilentlyContinue
foreach ($c in $listen) { if ((Test-CcmProcess $c.OwningProcess) -and ($ids -notcontains $c.OwningProcess)) { $ids+=$c.OwningProcess } }
foreach ($id in $ids) {
 Stop-Process -Id $id -ErrorAction SilentlyContinue
 Wait-Process -Id $id -Timeout 10 -ErrorAction SilentlyContinue
}
for ($i=0; $i -lt 20; $i++) {
 if (!(Get-NetTCPConnection -LocalPort 3443 -State Listen -ErrorAction SilentlyContinue)) { break }
 Start-Sleep -Milliseconds 250
}
