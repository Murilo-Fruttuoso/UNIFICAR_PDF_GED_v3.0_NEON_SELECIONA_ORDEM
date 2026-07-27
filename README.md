# Unificador de PDFs / Notas de Débito

## Arquivos

- `unificar_pdfs_v3.0_Seleciona_Ordem.py` — versão original (mantida como referência/histórico).
- `unificar_pdfs_v4.0_ND.py` — **nova versão**, com as melhorias solicitadas.

## O que mudou na v4.0

1. **Junta também o arquivo da nota de débito.**
   O script agora reconhece automaticamente arquivos cujo nome comece com
   `"Nota de débito - Neon Pagamentos ND"` (ex: `Nota de débito - Neon Pagamentos ND 1262.pdf`)
   e os insere como a **primeira página** do PDF unificado correspondente.

2. **Compressão do arquivo final.**
   Antes de salvar, o script recomprime as imagens internas do PDF (reduz
   resolução/qualidade quando necessário) e os streams de conteúdo, além de
   aplicar uma segunda passada de otimização via `pikepdf` (quando disponível
   no ambiente). Isso reduz bastante o tamanho final sem comprometer a
   legibilidade dos documentos.

3. **Nova coluna "ND" na base de nomenclatura.**
   A base (Excel ou CSV) agora deve ter 3 colunas:

   | Coluna | Conteúdo |
   |---|---|
   | A | Código do arquivo (ex: `100`) |
   | B | Nome de saída (usado só quando o código **não** tiver ND definido) |
   | C | **ND** — número da nota de débito à qual aquele código pertence |

   Para cada número de ND presente na coluna C, o script gera **um único PDF
   unificado** contendo: o arquivo da nota de débito daquele ND + todos os
   comprovantes de todos os códigos vinculados a ele (respeitando a ordem em
   que aparecem na base e a ordem de letras informada pelo usuário).

   O arquivo final é nomeado como:
   ```
   Nota de débito - Neon Pagamentos ND <número>-<ano atual>.pdf
   ```
   Exemplo: `Nota de débito - Neon Pagamentos ND 1262-2026.pdf`
   (o ano é sempre o ano atual no momento em que o script é executado).

   Códigos que **não** tiverem ND preenchido na base continuam sendo
   processados individualmente, exatamente como na v3.0 (nome vindo da
   coluna B).

4. **Log de processamento.**
   Ao final da execução, é criado um arquivo
   `LOG_processamento_AAAAMMDD_HHMMSS.txt` na pasta de saída, contendo:
   - Resumo geral (quantos PDFs foram lidos, gerados, ignorados, avisos).
   - Lista de todos os PDFs gerados, com o ND correspondente e os arquivos
     que foram unidos em cada um.
   - Lista de avisos (ex: código sem arquivo correspondente, ND sem nota
     encontrada, código duplicado entre NDs, etc.).
   - Lista de arquivos possivelmente ignorados, com o motivo (ex: nome fora
     do padrão esperado, nota de débito cujo ND não consta na base, nota
     duplicada para o mesmo ND).

## Executável (.exe) — rodar em qualquer máquina Windows sem instalar Python

Já existe um executável pronto em `dist/UnificarPDFs_ND.exe`. Basta copiar
esse único arquivo para qualquer computador Windows (32 ou 64 bits) e dar
duplo clique — não é necessário instalar Python nem nenhuma dependência.

Ele abre as mesmas janelas gráficas do script (pasta de entrada, pasta de
saída, base de nomenclatura e, se necessário, a ordem das letras) e gera os
PDFs unificados + o log, exatamente como a versão Python.

Para gerar o `.exe` novamente após alterar o script, veja as instruções em
`build_exe/COMO_GERAR_O_EXECUTAVEL.md`.

## Como executar via Python (alternativa ao .exe)

```bash
python3 unificar_pdfs_v4.0_ND.py
```

O script pedirá, em janelas gráficas:
1. Pasta de entrada (onde estão os comprovantes e os arquivos de nota).
2. Pasta de saída (onde serão salvos os PDFs unificados e o log).
3. Arquivo de base (Excel/CSV com as colunas Código / Nome / ND).
4. Caso existam comprovantes com letras no nome (ex: `100a.pdf`), será
   perguntada a ordem de leitura das letras (ex: `C,G,A`).

## Dependências

Ver `requirements.txt`. Instalação:

```bash
pip install -r requirements.txt
```

> Observação: `pikepdf` é opcional — se não estiver instalado, o script
> ainda funciona e comprime as imagens/streams via `pypdf`/`Pillow`, apenas
> sem a passada extra de otimização de objetos do PDF.
