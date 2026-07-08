@echo off
setlocal

set "APPDIR=%LOCALAPPDATA%\ConversorPDF"
set "EXE=%APPDIR%\ConversorPDF.exe"
set "SRC=%~dp0ConversorPDF"

echo Instalando "Converter em PDF"...
echo.

if not exist "%SRC%\ConversorPDF.exe" (
    echo ERRO: pasta ConversorPDF nao encontrada aqui.
    echo Certifique-se de que install.bat esta ao lado da pasta ConversorPDF.
    pause
    exit /b 1
)

if not exist "%APPDIR%" mkdir "%APPDIR%"
robocopy "%SRC%" "%APPDIR%" /E /NFL /NDL /NJH /NJS >nul
if %errorlevel% GEQ 8 (
    echo ERRO ao copiar os arquivos do programa.
    pause
    exit /b 1
)

reg add "HKCU\Software\Classes\*\shell\ConverterParaPDF" /ve /d "Converter em PDF" /f >nul
reg add "HKCU\Software\Classes\*\shell\ConverterParaPDF" /v "Icon" /d "%SystemRoot%\System32\imageres.dll,-102" /f >nul
reg delete "HKCU\Software\Classes\*\shell\ConverterParaPDF" /v "MultiSelectModel" /f >nul 2>&1
reg add "HKCU\Software\Classes\*\shell\ConverterParaPDF\command" /ve /d "\"%EXE%\" %%1" /f >nul

if errorlevel 1 (
    echo ERRO ao gravar as chaves de registro.
    pause
    exit /b 1
)

echo Instalado com sucesso!
echo.
echo Como usar:
echo  1. Selecione um ou mais arquivos no Explorer (imagens, Word, Excel, PowerPoint ou txt)
echo  2. Clique com o botao direito
echo  3. No Windows 11, clique em "Mostrar mais opcoes" primeiro
echo  4. Clique em "Converter em PDF"
echo.
pause
