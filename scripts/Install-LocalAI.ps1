param([string]$Model='qwen3:4b')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
if ($Model -notin @('qwen3:4b','qwen3:1.7b')) { throw 'Chi ho tro mo hinh local qwen3:4b / qwen3:1.7b.' }
$target=Join-Path $root 'runtime/ollama'
$exe=Join-Path $target 'ollama.exe'
if (!(Test-Path -LiteralPath $exe)) {
 $download=Join-Path $root 'runtime/ollama-download'
 New-Item -ItemType Directory -Force -Path $download | Out-Null
 $zip=Join-Path $download 'ollama-cpu.zip'
 # SHA256 of CPU ZIP byte range verified from official v0.35.0 HTTPS release.
 $expected='a4aa5ae1d0edcc89c0d07ddb80d0ba4f87586a543f5391bc7055906a30383aea'
 if (!(Test-Path -LiteralPath $zip) -or (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLower() -ne $expected) {
  Write-Host 'Tai Ollama CPU chinh thuc v0.35.0 (~36 MB). Khong gui ho so len mang.'
  & curl.exe -L --fail --retry 2 --range '1424840914-1461196157' --output $zip 'https://github.com/ollama/ollama/releases/download/v0.35.0/ollama-windows-amd64.zip'
  if ($LASTEXITCODE -ne 0) { throw 'Tai Ollama that bai; chay lai khi mang on dinh.' }
 }
 if ((Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLower() -ne $expected) { throw 'Checksum Ollama khong khop. Khong cai runtime.' }
 $python=Join-Path $root 'runtime/python/python.exe'
 & $python (Join-Path $PSScriptRoot 'extract_ollama_cpu.py') $zip $target
 if ($LASTEXITCODE -ne 0) { throw 'Giai nen Ollama that bai.' }
 $signature=Get-AuthenticodeSignature -LiteralPath $exe
 if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notlike '*Ollama Inc.*') { throw 'Chu ky Ollama khong hop le.' }
}
& (Join-Path $PSScriptRoot 'Start-LocalAI.ps1')
$env:OLLAMA_HOST='127.0.0.1:11434'
Write-Host ('Tai mo hinh '+$Model+'. Chi tai trong luong mo hinh; khong gui ho so.')
& $exe pull $Model
if ($LASTEXITCODE -ne 0) { throw 'Tai mo hinh chua xong. Chay lai de tiep tuc.' }
$s=@{model=$Model;enabled=$true} | ConvertTo-Json
[IO.File]::WriteAllText((Join-Path $root 'config/ai.json'),$s,(New-Object Text.UTF8Encoding($false)))
Write-Host 'Da cai AI cuc bo. Mo CCM-AI > AI cuc bo > Thu AI bang du lieu minh hoa.'
