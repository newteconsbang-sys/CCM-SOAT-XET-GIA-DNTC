$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$pidFile=Join-Path $root 'data/server.pid'
if (Test-Path -LiteralPath $pidFile) {
 $ccmPid=[int](Get-Content -LiteralPath $pidFile)
 $process=Get-CimInstance Win32_Process -Filter "ProcessId=$ccmPid"
 if ($process -and $process.CommandLine.Replace('/','\').Contains($root+'\backend\run.py')) {
  Stop-Process -Id $ccmPid
  Wait-Process -Id $ccmPid -Timeout 10 -ErrorAction SilentlyContinue
 }
}
