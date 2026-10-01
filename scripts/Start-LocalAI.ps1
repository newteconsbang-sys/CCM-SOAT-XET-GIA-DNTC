$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$exe=Join-Path $root 'runtime/ollama/ollama.exe'
if (!(Test-Path -LiteralPath $exe)) { throw 'Chua cai Ollama. Chay Cai AI cuc bo.cmd.' }
try { $null=Invoke-RestMethod 'http://127.0.0.1:11434/api/tags' -TimeoutSec 3; return } catch {}
$env:OLLAMA_HOST='127.0.0.1:11434'
$env:OLLAMA_MODELS=Join-Path $root 'data/models'
$env:OLLAMA_NO_CLOUD='1'
$env:OLLAMA_CONTEXT_LENGTH='8192'
$env:CUDA_VISIBLE_DEVICES='-1'
New-Item -ItemType Directory -Force -Path $env:OLLAMA_MODELS | Out-Null
$p=Start-Process -FilePath $exe -ArgumentList 'serve' -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $root 'data/ollama-out.log') -RedirectStandardError (Join-Path $root 'data/ollama-error.log')
$p.Id | Set-Content -LiteralPath (Join-Path $root 'data/ollama.pid')
for ($i=0;$i -lt 30;$i++) {
 Start-Sleep -Milliseconds 500
 try { $null=Invoke-RestMethod 'http://127.0.0.1:11434/api/tags' -TimeoutSec 2; return } catch {}
}
throw 'Ollama chua san sang. Xem data/ollama-error.log.'
