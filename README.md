# Unificador de PDFs / Notas de Débito

## Arquivos

- `unificar_pdfs_v3.0_Seleciona_Ordem.py` — versão original (mantida como referência/histórico).
- `unificar_pdfs_v4.0_ND.py` — versão anterior com agrupamento por ND, compressão e log
  (mantida como referência/histórico — apresentava os 3 problemas corrigidos na v5.0, veja abaixo).
- `unificar_pdfs_v5.0_ND.py` — **versão atual, recomendada.**

O script lê uma base `.xlsx` com `ID`, `LANÇAMENTO`, `DATA DO PAGAMENTO` e `ND`, solicita as pastas por janelas e gera um log Excel estruturado.

## O que mudou na v5.0 (correções sobre a v4.0)

A v5.0 corrige 3 problemas identificados no uso real da v4.0:

### 1. Nova base de nomenclatura, com data de pagamento, e ordenação por letra + data entre códigos diferentes

A base (Excel ou CSV) agora pode ter **cabeçalho**, com as colunas (em
qualquer ordem):

| Coluna | Conteúdo |
|---|---|
| `Id` | Código do arquivo (ex: `10`) |
| `LANÇAMENTO` | Nome/descrição (usado só quando o código **não** tiver ND definido) |
| `DATA DO PAGAMENTO` | Data usada para ordenar os códigos dentro de um mesmo ND |
| `ND` | Número da nota de débito à qual aquele código pertence |

Para cada número de ND presente na coluna `ND`, o script junta **todos os
comprovantes de todos os códigos daquele ND em uma única lista** e ordena
essa lista inteira por:

1. **Posição da letra do arquivo** (ex: `10a.pdf`, `20a.pdf`) na ordem que o
   usuário informar no popup (ex: `A,B`) — critério principal;
2. **Data de pagamento** do código a que o arquivo pertence, da mais antiga
   para a mais recente — dentro de arquivos com a mesma letra, arquivos de
   códigos diferentes ficam entrelaçados por data;
3. Ordem da linha na planilha (desempate quando datas são iguais/ausentes,
   respeitando a ordem em que os códigos foram colocados na base);
4. Número extra no nome do arquivo e nome do arquivo (desempates finais).

Exemplo: código `10` (pago em 15/01) e código `20` (pago em 10/01) pertencem
ao mesmo ND. Com ordem de letras `A,B`, o resultado é:
`20a.pdf, 10a.pdf, 10b.pdf` — todas as páginas "A" antes das páginas "B", e
dentro de cada letra, o código pago primeiro (20, em 10/01) vem antes do
código pago depois (10, em 15/01).

A base **antiga** (sem cabeçalho: coluna A = código, B = nome, C = ND, sem
data de pagamento) continua funcionando exatamente como antes, para manter
compatibilidade com planilhas já existentes — nesse caso a ordenação usa
apenas letra + ordem da planilha (sem data, pois ela não existe nesse
formato).

A data de pagamento é aceita em qualquer um destes formatos: célula já
formatada como data no Excel, número serial do Excel (quando a célula não
está formatada como data), ou texto `dd/mm/aaaa`, `dd-mm-aaaa`, `aaaa-mm-dd`,
`dd/mm/aa` ou `dd.mm.aaaa`.

> **Sobre arquivos `.csv`:** o script detecta automaticamente a codificação
> do arquivo (UTF-8, UTF-16 ou `cp1252`/Windows-1252 — este último é o
> padrão quando o CSV é salvo pelo Excel em português) e também o separador
> usado (`,`, `;` ou tabulação — o Excel em português normalmente usa `;`).
> Não é necessário se preocupar com esses detalhes ao exportar a planilha
> do Excel como CSV.

### 2. Log de processamento XLSX sempre gerado, mesmo em caso de erro

Na v4.0, se qualquer erro inesperado ocorresse durante o processamento (por
exemplo, um PDF corrompido, uma planilha com formato inesperado, etc.), o
script podia ser interrompido **antes** de gerar o log — e como o `.exe` é
gerado sem janela de console (`--noconsole`), o usuário não via nem o erro
nem tinha um log para diagnosticar o problema.

Na v5.0, toda a lógica principal roda dentro de um bloco protegido: se
ocorrer qualquer erro fatal, o script:
- Captura o erro completo (traceback);
- **Ainda assim gera o arquivo de log XLSX** na pasta de saída, com a aba `Resumo`
  indicando `ERRO FATAL` e contendo o detalhamento técnico do problema;
- Mostra uma mensagem na tela avisando que houve erro e indicando o caminho
  exato do log gerado.

Ou seja: **o log é sempre criado**, tanto em execuções bem-sucedidas quanto
em execuções que falharam, e também nos casos em que a operação é
cancelada pelo usuário ou nenhum PDF válido é encontrado na pasta.

O log XLSX contém as abas `Resumo`, `Arquivos gerados`, `Avisos` e `Ignorados`, com:
- Resumo geral (quantos PDFs foram lidos, gerados, ignorados, avisos);
- Lista de todos os PDFs gerados, com o ND correspondente e os arquivos que
  foram unidos em cada um (na ordem exata em que entraram no PDF);
- Lista de avisos (ex: código sem arquivo correspondente, ND sem nota
  encontrada, código duplicado entre NDs, etc.);
- Lista de arquivos possivelmente ignorados, com o motivo (ex: nome fora do
  padrão esperado, nota de débito cujo ND não consta na base, nota
  duplicada para o mesmo ND).

### 3. Reconhecimento de qualquer "Nota de débito", independente da marca

Na v4.0, o script só reconhecia arquivos cujo nome começasse **exatamente**
com `"Nota de débito - Neon Pagamentos ND"`. Isso fazia com que notas de
outras marcas/parceiros (ex: `"Nota de débito - Neon Consiga+ ND 1266.pdf"`)
fossem ignoradas.

Na v5.0, o script reconhece **qualquer** arquivo cujo nome comece com
`"Nota de débito"` (sem acentuação e maiúsculas/minúsculas fazem diferença),
independentemente do que vem depois — a marca/texto até a palavra `ND` é
extraída automaticamente do próprio nome do arquivo e preservada no nome de
saída.

O arquivo final é nomeado como:
```
Nota de débito - Neon Pagamentos ND <número>.pdf
```

Exemplos:
- ND 1266 → gera `Nota de débito - Neon Pagamentos ND 1266.pdf`
- ND 1267 → gera `Nota de débito - Neon Pagamentos ND 1267.pdf`
- ND 1262 → gera `Nota de débito - Neon Pagamentos ND 1262.pdf`

**Importante:** o número usado no nome final (`1266`, `1267`, `1262`, etc.)
é sempre o valor da coluna `ND` da planilha de base, **não** o número que
aparece no nome do próprio arquivo de nota — isso evita divergência caso o
nome do arquivo da nota e a planilha estejam com números diferentes por
algum erro de digitação.

## Funcionalidades mantidas da v4.0

1. **Junta o arquivo da nota de débito como primeira página** do PDF
   unificado de cada ND.
2. **Compressão do arquivo final** — recomprime imagens internas (reduz
   resolução/qualidade quando necessário) e streams de conteúdo, além de uma
   segunda passada de otimização via `pikepdf` (quando disponível). Reduz
   bastante o tamanho final sem comprometer a legibilidade dos documentos.
3. Códigos que **não** tiverem ND preenchido na base continuam sendo
   processados individualmente (nome vindo da coluna de nome/lançamento).

## Executável (.exe) — rodar em qualquer máquina Windows sem instalar Python

Já existe um executável pronto em `dist/UnificarPDFs_ND.exe`, gerado a partir
da v5.0. Basta copiar esse único arquivo para qualquer computador Windows
(32 ou 64 bits) e dar duplo clique — não é necessário instalar Python nem
nenhuma dependência.

Ele abre as mesmas janelas gráficas do script (pasta de entrada, pasta de
saída, base de nomenclatura e, se necessário, a ordem das letras) e gera os
PDFs unificados + o log, exatamente como a versão Python.

Para gerar o `.exe` novamente após alterar o script, veja as instruções em
`build_exe/COMO_GERAR_O_EXECUTAVEL.md`.

## Como executar via Python (alternativa ao .exe)

```bash
python3 unificar_pdfs_v5.0_ND.py
```

O script pedirá, em janelas gráficas:
1. Pasta de entrada (onde estão os comprovantes e os arquivos de nota).
2. Pasta de saída (onde serão salvos os PDFs unificados e o log).
3. Arquivo de base (Excel/CSV com as colunas `Id` / `LANÇAMENTO` /
   `DATA DO PAGAMENTO` / `ND`, ou o formato antigo sem cabeçalho).
4. Caso existam comprovantes com letras no nome (ex: `10a.pdf`), será
   perguntada a ordem de leitura das letras (ex: `C,G,A`).

Ao final, sempre é exibida uma mensagem indicando quantos PDFs foram
gerados e o caminho do arquivo de log — mesmo se algo tiver dado errado.

## Dependências

Ver `requirements.txt`. Instalação:

```bash
pip install -r requirements.txt
```

> Observação: `pikepdf` é opcional — se não estiver instalado, o script
> ainda funciona e comprime as imagens/streams via `pypdf`/`Pillow`, apenas
> sem a passada extra de otimização de objetos do PDF.
