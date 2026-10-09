# AssTI 2.0

Aplicativo desktop local para Windows 11, com duas abas independentes:

- **PDF / Chaves NFCom:** importa PDFs, pastas (incluindo subpastas) e ZIPs, identifica conta e chaves, valida, copia e exporta XLSX.
- **XML:** mantém a interface original de importação, histórico e configurações. `core.py` está inalterado, incluindo cálculos financeiros, classificações e os três modelos de planilha.

## Arquitetura

`main.py` usa PySide6/Qt e envolve a interface XML original numa segunda aba, sem reconstruir seus controles. `core.py` usa defusedxml, Decimal e openpyxl. `pdf_core.py` concentra leitura local com PyMuPDF, validação, ZIPs, processamento e exportação; `pdf_ui.py` implementa a primeira aba com QThread, sinais de progresso e cancelamento cooperativo.

O checkout original contém apenas os quatro arquivos Python/dependências: não inclui `assets/`, `version_info.txt`, atalhos ou documentos auxiliares referenciados pelo compilador. Recursos existentes são utilizados se fornecidos; sua ausência não impede compilar.

## Executar no Windows

Use Python 3.11 ou posterior (64 bits), em uma pasta com permissão de escrita. No PowerShell:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

Não são necessárias credenciais, Excel instalado ou serviços externos para processar documentos. Configurações, histórico XML e log ficam no diretório de dados local do aplicativo.

## Usar a aba PDF

1. Adicione PDFs, selecione uma pasta ou importe ZIPs. A lista evita importar o mesmo caminho/entrada de ZIP repetidamente.
2. Ative OCR se necessário e configurado; clique em **Iniciar extração**. Um novo processamento substitui os resultados anteriores, mantendo a lista de arquivos.
3. Confira status e observações. Arquivos individuais ilegíveis não interrompem o lote. É possível cancelar; os documentos concluídos permanecem disponíveis. Durante o processamento, aguarde ou cancele antes de fechar.
4. Use Ctrl/Shift para selecionar linhas. **Copiar selecionadas** copia as chaves das linhas escolhidas, inclusive inválidas, conforme a seleção explícita. **Copiar todas as chaves válidas** exclui as inválidas. Cada chave ocupa uma linha, sem separadores. Ocorrências duplicadas permanecem visíveis e são identificadas.
5. Exporte o Excel ou limpe os resultados e arquivos para um novo lote.

A planilha contém as seis colunas solicitadas e uma aba `Resumo`. Conta, chaves e nomes são texto literal, preservando zeros e evitando interpretação como fórmula. Contadores de chaves contam ocorrências; duplicadas são ocorrências posteriores à primeira, mesmo dentro do mesmo PDF. O resumo conta documentos processados, não documentos pendentes/cancelados.

## Extração e validação

Chaves são procuradas junto ao rótulo `Chave de acesso` ou em contexto NFCom com modelo 62. Aceita espaços, pontos, hífens e quebras de linha; não completa nem corrige dígitos. Valida 44 dígitos ASCII, código de UF, mês, CNPJ numérico e seus verificadores, modelo 62, número da nota, tipo de emissão e DV módulo 11. Isso verifica a estrutura, não autorização na SEFAZ, assinatura ou situação fiscal. Chaves com quantidade inválida junto ao rótulo são marcadas para revisão.

Não consulta servidores externos, QR codes ou portais fiscais. Contas sem rótulo reconhecível permanecem vazias, com observação. Layouts não reconhecidos ficam para revisão. Limites defensivos: 100 MB por PDF e 2 GB de PDFs descompactados por ZIP; não há limite fixo de 198 documentos. ZIPs são lidos diretamente, sem gravar seus caminhos no disco.

## OCR opcional

O fluxo normal funciona sem OCR. PyMuPDF usa Tesseract local em páginas completamente sem texto. Instale/disponibilize os modelos `por.traineddata` e `eng.traineddata` e configure `TESSDATA_PREFIX` para a pasta que os contém. No Windows, modelos podem ficar numa pasta `tessdata` junto à distribuição:

```powershell
$env:TESSDATA_PREFIX = "$PWD\tessdata"
.\AssTI.exe
```

Os modelos oficiais estão em https://github.com/tesseract-ocr/tessdata_fast. A pasta deve existir no computador destinatário; os modelos não são baixados automaticamente pelo aplicativo ou incluídos no ZIP de distribuição. OCR indisponível produz observação de revisão e não interrompe o lote. Páginas com texto parcial ou chave em imagem ao lado de texto exigem revisão manual; OCR pode produzir erros, portanto confira o original.

## Testes

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Foram executados **15 testes com sucesso** no ambiente Linux, com Qt offscreen e OCR português/inglês. A suíte cobre lote de 198 PDFs em ZIP, pasta, chaves formatadas, múltiplas chaves, duplicadas, DV/estrutura inválidos, códigos de pagamento ignorados, PDF sem chave, ilegível/protegido, limites de ZIP, cancelamento, OCR real e indisponível, inicialização das duas abas, preservação de estado, clipboard e Excel, classificação/somas XML e contratos de exportação. Sem `TESSDATA_PREFIX`, o teste de OCR real é explicitamente pulado; os outros continuam.

Fixtures são fictícias. **Ainda não foi validado com faturas reais representativas da Claro.** A suíte não substitui essa validação, nem os testes finais de distribuição em Windows 11.

## Gerar o executável portátil

No Windows, depois de instalar dependências:

```powershell
.\.venv\Scripts\python.exe build_windows.py
```

O script executa a suíte antes do PyInstaller, inclui PyMuPDF e seus recursos nativos, mantém ícones/recursos opcionais quando presentes, copia fontes e licenças, e gera `dist/AssTI/AssTI.exe` e `AssTI_Compartilhar.zip`. Distribua a pasta inteira, mantendo `_internal` junto do executável; o destinatário não precisa instalar Python. Ícones personalizados: `assets/assti.ico` e `assets/assti.png`. O build Windows não faz compilação cruzada: **o `.exe` não foi gerado nem testado neste ambiente Linux**.

Confira as licenças das dependências antes de distribuir: PyMuPDF é disponibilizado sob AGPL ou licença comercial, e o script copia as licenças encontradas para `Licencas`.

## Ambiente Codex

Ambiente preparado em `/workspace/assti-venv`; execute:

```bash
cd /workspace/assTI
TESSDATA_PREFIX=/workspace/assti-tessdata QT_QPA_PLATFORM=offscreen XDG_CACHE_HOME=/tmp/assti-cache /workspace/assti-venv/bin/python -m unittest discover -s tests -v
```

Qt offscreen permite validar widgets e clipboard internos sem monitor; não verifica a aparência ou o clipboard do Windows. Cada tarefa já é isolada: utilize este checkout, sem criar worktrees. Alterações estão na branch `feature/assti-2-extrator-nfcom` para revisão antes da integração em `main`.
