@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==== Day CCM-AI Soat xet gia len GitHub ====
where git >nul 2>nul && goto :hasgit
if exist "%ProgramFiles%\Git\cmd\git.exe" goto :addpath
echo Chua co Git tren may - dang tu cai Git (mien phi, chinh hang tu winget)...
winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
if not exist "%ProgramFiles%\Git\cmd\git.exe" (
  echo.
  echo Khong tu cai duoc Git. Hay tai va cai tai: https://git-scm.com/download/win
  echo ^(bam Next lien tuc la xong^), sau do chay lai file nay.
  start "" https://git-scm.com/download/win
  pause & exit /b 1
)
:addpath
set "PATH=%ProgramFiles%\Git\cmd;%PATH%"
:hasgit
set /p VER=<backend\VERSION
git config user.name >nul 2>nul || git config user.name "newteconsbang-sys"
git config user.email >nul 2>nul || git config user.email "newteconsbang-sys@users.noreply.github.com"
git remote get-url origin >nul 2>nul && git remote set-url origin https://github.com/newteconsbang-sys/CCM-SOAT-XET-GIA-DNTC.git
git remote get-url origin >nul 2>nul || git remote add origin https://github.com/newteconsbang-sys/CCM-SOAT-XET-GIA-DNTC.git
rem Bo theo doi (KHONG xoa file) cac thu muc du lieu that neu lo da them truoc day
git rm -r --cached -q --ignore-unmatch data uploads backups reports config runtime .venv .pnpm-store frontend/node_modules frontend/dist >nul 2>nul
git add -A
git commit -q -m "CCM-AI Soat xet gia - phien ban %VER%" || echo (Khong co thay doi moi de commit)
git branch -M main
git fetch origin >nul 2>nul
git rev-parse --verify -q origin/main >nul && git pull --no-rebase --no-edit --allow-unrelated-histories -X ours origin main
git push -u origin main
if errorlevel 1 (echo. & echo DAY LEN THAT BAI - xem thong bao loi o tren. & pause & exit /b 1)
echo.
echo DA DAY LEN: https://github.com/newteconsbang-sys/CCM-SOAT-XET-GIA-DNTC  (phien ban %VER%)
pause
