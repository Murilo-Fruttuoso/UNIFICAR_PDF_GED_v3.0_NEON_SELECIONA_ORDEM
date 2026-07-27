# Como gerar o executável (.exe) a partir do script Python

O executável já compilado está em `dist/UnificarPDFs_ND.exe` (pronto para
uso em qualquer Windows 32 ou 64 bits, sem precisar instalar Python).

Caso precise gerar novamente (ex.: após alterar o script), siga os passos
abaixo em uma máquina Windows (ou usando Wine em Linux, como foi feito
neste ambiente):

## 1. Instalar o Python (Windows)
Baixe e instale o Python 3.11+ em https://www.python.org/downloads/
(marque a opção "Add python.exe to PATH" durante a instalação).

## 2. Instalar as dependências
Abra o "Prompt de Comando" (cmd) e execute:

```
pip install pypdf openpyxl Pillow pyinstaller
```

> `pikepdf` é opcional (melhora um pouco mais a compressão). Se quiser
> incluir, instale com `pip install pikepdf` antes do passo 3 — o script
> detecta automaticamente se está disponível.

## 3. Gerar o executável

Copie o arquivo `unificar_pdfs_v5.0_ND.py` (renomeie para `unificar_pdfs.py`,
ou ajuste o comando abaixo) para uma pasta e execute, dentro dessa pasta:

```
pyinstaller --onefile --noconsole --name "UnificarPDFs_ND" --clean unificar_pdfs_v5.0_ND.py
```

O executável final será gerado em `dist\UnificarPDFs_ND.exe`.

- `--onefile`: gera um único arquivo `.exe` (mais fácil de distribuir).
- `--noconsole`: não abre uma janela de terminal preta ao rodar (o
  programa já usa janelas gráficas próprias para pedir pastas/base/ordem).

## 4. Distribuir

Basta copiar o arquivo `UnificarPDFs_ND.exe` para qualquer máquina Windows
— não é necessário ter Python instalado nela. Ao clicar duas vezes, o
programa abrirá as janelas para escolher a pasta de entrada, pasta de
saída e a base de nomenclatura, exatamente como o script Python original.

## Observações sobre este ambiente de build

Este executável foi compilado usando Wine (emulador de Windows em Linux)
dentro do sandbox, com Python 3.11.9 (build Windows oficial) + PyInstaller
6.21.0, arquitetura **win32** (compatível com Windows 32 e 64 bits). Os
testes de fumaça (imports de `tkinter`, `pypdf`, `openpyxl`, `PIL`, e a
execução completa da lógica de agrupamento por ND/data de pagamento,
reconhecimento genérico de "Nota de débito" e geração garantida do log)
foram validados com sucesso dentro do próprio ambiente Wine antes da
entrega do binário (v5.0 — recompilado para corrigir os 3 problemas
relatados sobre a versão anterior: ordenação cross-código por data,
log sempre gerado mesmo em erro fatal, e reconhecimento de notas de
débito de qualquer marca).
