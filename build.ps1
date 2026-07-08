# Gera a pasta standalone ConversorPDF\ (nao precisa Python na maquina de destino).
# Roda so na maquina de build (Maike). Requer Python instalado aqui.
#
# --onedir (nao --onefile): onefile reextrai o pacote inteiro pra uma pasta temp
# TODA vez que abre, o que e lento demais em HD (nao SSD) quando o Explorer
# dispara varias instancias quase simultaneas (uma por arquivo selecionado).
# onedir ja fica extraido em disco, entao cada instancia abre rapido.

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

Write-Host "Instalando PyInstaller (ambiente de build)..."
py -m pip install --upgrade pyinstaller | Out-Null

Write-Host "Gerando pasta ConversorPDF\..."
py -m PyInstaller --onedir --noconsole --name ConversorPDF `
    --hidden-import win32timezone `
    --clean `
    convert_to_pdf.py

Write-Host ""
Write-Host "Pronto: $here\dist\ConversorPDF\"
Write-Host "Teste essa pasta isolada antes de empacotar (idealmente numa maquina/pasta sem o Python de dev no PATH)."
Write-Host "O pacote de distribuicao precisa da pasta ConversorPDF\ inteira, nao so do .exe."
