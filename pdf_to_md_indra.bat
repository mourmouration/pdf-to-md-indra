@echo off
chcp 65001 >nul
setlocal

set /p FOLDERPATH="Введите путь к папке: "

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0pdf_to_md_indra.ps1" -FolderPath "%FOLDERPATH%"

echo.
pause
