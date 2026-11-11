@echo off
rem Tat dich vu CCM-AI dang chay roi bat lai ban moi nhat va mo web.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Stop-CCM.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Open-Web.ps1"
if errorlevel 1 (
  echo Khong khoi dong duoc. Xem data\startup-error.log va data\server-error.log
  pause
)
