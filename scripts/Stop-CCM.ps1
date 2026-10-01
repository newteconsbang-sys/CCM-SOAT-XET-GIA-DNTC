$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$pidFile=Join-Path $root 'data/server.pid'
if (Test-Path -LiteralPath $pidFile) {
 $ccmPid=[int](Get-Content -LiteralPath $pidFile)
 $process=Get-CimInstance Win32_Process -Filter "ProcessId=$ccmPid"
 if ($process -and $process.CommandLine -like ('*'+$root+'\backend/run.py*')) { Stop-Process -Id $ccmPid }
}
