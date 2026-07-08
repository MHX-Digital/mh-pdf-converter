# mh-pdf-converter

Ferramenta de linha de comando para Windows que adiciona a opção **"Converter em PDF"** ao menu de contexto do Explorer. Selecione dois ou mais arquivos (imagens, documentos do Office ou texto), clique com o botão direito e junte tudo em um único PDF, na ordem que você escolher.

## O que ela faz

- Converte cada arquivo selecionado para PDF internamente e depois junta tudo em um único documento.
- Suporta imagens (`.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.tiff`, `.webp`), Word (`.doc`, `.docx`, `.rtf`), Excel (`.xls`, `.xlsx`, `.csv`), PowerPoint (`.ppt`, `.pptx`) e texto puro (`.txt`).
- Antes de gerar o PDF, mostra uma tela para reordenar os arquivos manualmente.
- Pergunta o nome e o local do arquivo final através de uma janela padrão "Salvar como".
- Roda como executável standalone: quem instala não precisa ter Python na máquina.
- Instalação por usuário (`HKEY_CURRENT_USER`), sem exigir privilégios de administrador.

## Como funciona por baixo dos panos

O Windows Explorer, pelo menos nas versões testadas, não respeita o valor de registro `MultiSelectModel=Player`, que na teoria deveria fazer o shell invocar o programa uma única vez com todos os arquivos selecionados. Na prática, ele dispara uma instância separada do executável para cada arquivo escolhido.

Para contornar isso, o programa usa um esquema de coordenação entre processos: a primeira instância a conseguir um mutex nomeado do Windows vira a "líder" e espera as instâncias irmãs reportarem seus arquivos através de uma fila salva em um arquivo temporário. As demais instâncias ("seguidoras") apenas anexam seu arquivo à fila e encerram sem abrir nenhuma interface. A espera da líder é adaptativa (entre 1,5 e 6 segundos), parando assim que a fila fica um tempo sem crescer.

O executável é empacotado no modo `--onedir` do PyInstaller, e não `--onefile`. O modo onefile reextrai o pacote inteiro para uma pasta temporária a cada execução. Isso se mostrou lento demais em discos mecânicos quando dez ou mais processos abrem quase ao mesmo tempo, a ponto de quebrar a coordenação descrita acima.

A conversão de arquivos do Office usa automação COM (`win32com`) com Word, Excel e PowerPoint já instalados na máquina. Um detalhe que custou tempo para descobrir: o Excel depende do serviço **Spooler de Impressão do Windows** estar ativo para exportar PDF via COM, mesmo sem nenhuma impressora física conectada. Se esse serviço estiver desabilitado, a conversão de planilhas falha com um erro do próprio Excel.

## Requisitos

- Windows 10 ou 11.
- Microsoft Office instalado (Word, Excel e/ou PowerPoint), apenas para os tipos de arquivo correspondentes. Imagens e texto funcionam sem Office.
- Serviço Spooler de Impressão do Windows ativo, para conversão de planilhas Excel.

## Instalação (uso final)

1. Baixe o pacote da última [release](../../releases).
2. Descompacte em qualquer pasta.
3. Execute `install.bat`. Não é necessário ser administrador.
4. Selecione arquivos no Explorer, clique com o botão direito (no Windows 11, primeiro em "Mostrar mais opções") e escolha "Converter em PDF".

Para remover, execute `uninstall.bat`.

## Build a partir do código-fonte

Requer Python 3.11 ou mais recente no Windows.

```powershell
pip install -r requirements.txt
.\build.ps1
```

O resultado fica em `dist\ConversorPDF\`. Essa pasta inteira, junto com `install.bat`, `uninstall.bat` e `LEIA-ME.txt`, forma o pacote de distribuição.

## Limitações conhecidas

- Testado apenas em Windows 10 e 11 com Office instalado localmente; não há suporte a Office Online ou LibreOffice.
- Se o serviço Spooler de Impressão estiver desativado, a conversão de arquivos Excel falha (Word e PowerPoint não dependem desse serviço).
- A ordem inicial da lista de arquivos, antes da tela de reordenação, depende da ordem em que o Explorer dispara os processos e não é totalmente previsível. Por isso existe a etapa de reordenação manual.

## Licença

Distribuído sob a licença MIT. Veja o arquivo [LICENSE](LICENSE).
