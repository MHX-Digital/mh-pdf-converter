@echo off
setlocal

reg delete "HKCU\Software\Classes\*\shell\ConverterParaPDF" /f >nul 2>&1
rd /s /q "%LOCALAPPDATA%\ConversorPDF" >nul 2>&1

echo "Converter em PDF" foi removido do menu de contexto.
pause
