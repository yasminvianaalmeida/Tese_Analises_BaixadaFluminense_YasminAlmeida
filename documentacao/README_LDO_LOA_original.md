# Análise das LDOs e LOAs Municipais (item 3.3 da tese)

Código usado na tese de doutorado de Yasmin Viana ("Geoinformação e gestão pública municipal na
Baixada Fluminense Histórica") para o item **3.3 — Análise das Leis de Diretrizes Orçamentárias
(LDO) e Leis Orçamentárias Anuais (LOA)**, que mede o **Grau de Integração Orçamentária** da
geoinformação nas peças orçamentárias dos 8 municípios da Baixada Fluminense Histórica.

O vocabulário, os pesos de força técnica, os temas, as listas de apoio (expressões de
falso-positivo, termos fuzzy) e a rubrica de classificação usados por este código são os mesmos
publicados nos Anexos A, B e C do `Texto-Tese-Modificado.docx` e ficam no arquivo
`Parametros_LDO_LOA.xlsx`. Esse mesmo arquivo de parâmetros também é usado na análise dos
instrumentos urbanísticos (pasta `../3.2_Instrumentos_Urbanisticos/`), pois os dois eixos
compartilham parte do vocabulário técnico.

## O que cada script faz

1. **`analise_ldo_loa.py`** — script principal. Lê os parâmetros de `Parametros_LDO_LOA.xlsx`, lê
   os PDFs das LDOs/LOAs organizados em pastas por município (com OCR via Tesseract nas páginas
   digitalizadas), busca o vocabulário técnico, classifica as ocorrências em dois eixos (força ×
   ancoragem), pondera, deduplica, calcula os indicadores e o Grau de Integração Orçamentária, e
   grava tudo em `Resultados_LDO_LOA.xlsx` (com fórmulas, valores monetários candidatos por trecho
   e agregação por projeto/município/ano).
2. **`Ferramentas_LDO_LOA.py`** — funções utilitárias usadas pelo script principal (leitura de
   parâmetros, normalização de texto, extração monetária, etc.).
3. **`gerar_apendices.py`** *(opcional)* — gera um `.docx` com o vocabulário completo e o livro de
   código usados na análise, a partir de `Parametros_LDO_LOA.xlsx`, com um carimbo SHA-256 do
   arquivo de parâmetros (garante que o texto do apêndice bate exatamente com os parâmetros
   efetivamente usados na rodada).
4. **`Harness_Validacao.py`** *(opcional)* — ferramenta de validação inter-avaliadores: sorteia uma
   amostra dos resultados para dois avaliadores humanos rotularem manualmente, e depois calcula
   kappa de Cohen e F1 entre os rótulos humanos e a classificação automática (semente fixa
   `SEMENTE = 42`, para reprodutibilidade da amostragem).

## Passo a passo para rodar (Windows/CMD)

Pasta de trabalho recomendada: `C:\analiseldoloa` — os caminhos já vêm configurados no código
apontando para essa pasta (ver "Área de alteração" no topo de `analise_ldo_loa.py`), então o mais
simples é reproduzir essa estrutura exata. Coloque TODO o conteúdo desta pasta em
`C:\analiseldoloa`.

**Estrutura esperada (`C:\analiseldoloa`):**

```
C:\analiseldoloa\
   analise_ldo_loa.py
   Ferramentas_LDO_LOA.py
   gerar_apendices.py          (opcional)
   Harness_Validacao.py        (opcional)
   Parametros_LDO_LOA.xlsx     (vocabulário e regras -- Anexos A, B, C da tese)
   requirements.txt
   BELFORD ROXO LOA E LDO\     (uma pasta por município; PDFs dentro)
      LDO - Belford Roxo - 2017.pdf
      LOA - Belford Roxo - 2017.pdf
      ...
   DUQUE DE CAXIAS LOA E LDO\
   ... (demais municípios)
```

Os resultados são gravados em `C:\analiseldoloa\Resultados_LDO_LOA.xlsx`.

**Passo 1 — Instalar o Python**

Baixe o Python 3.12 em python.org/downloads. No instalador, marque "Add python.exe to PATH".
Confira no CMD: `py --version`.

**Passo 2 — Instalar o Tesseract (OCR) com Português**

Instale o "Tesseract OCR" (build da UB Mannheim). Em "Additional language data", marque
"Portuguese". Caminho padrão esperado pelo código: `C:\Program Files\Tesseract-OCR` (o programa já
procura automaticamente nesse caminho e em `C:\Program Files (x86)\Tesseract-OCR`).

**Passo 3 — Instalar as bibliotecas (uma vez)**

```
cd /d C:\analiseldoloa
py -m pip install -r requirements.txt
```

**Passo 4 — Conferir os documentos**

Coloque as pastas municipais (com os PDFs de LDO e LOA) dentro de `C:\analiseldoloa`, no mesmo
padrão de nomes usado na tese (`<MUNICÍPIO> LOA E LDO\`). Confirme que `Parametros_LDO_LOA.xlsx`
está em `C:\analiseldoloa` e está **fechado** no Excel (o script não consegue ler o arquivo se ele
estiver aberto). Não é preciso editar o código: os caminhos já apontam para `C:\analiseldoloa`.

**Passo 5 — Rodar a análise**

```
cd /d C:\analiseldoloa
py analise_ldo_loa.py
```

A primeira execução com OCR é demorada (pode levar horas em documentos grandes e escaneados).
Deixe terminar até aparecer "Concluído". Resultado: `C:\analiseldoloa\Resultados_LDO_LOA.xlsx`.

**Passo 6 (opcional) — Gerar os apêndices a partir dos parâmetros**

```
py gerar_apendices.py Parametros_LDO_LOA.xlsx APENDICES.docx
```

Gera o vocabulário e o livro de código com carimbo SHA-256 (garante que o texto bate com os
parâmetros usados).

**Passo 7 (opcional) — Validação (kappa/F1)**

```
py Harness_Validacao.py amostrar Resultados_LDO_LOA.xlsx amostra.xlsx
```

Preencha à mão as colunas `rotulo_humano_1` e `rotulo_humano_2` em `amostra.xlsx`, depois:

```
py Harness_Validacao.py avaliar amostra.xlsx relatorio_validacao.xlsx
```

## Rodando fora do Windows (Linux/Mac)

O código foi escrito e documentado para Windows/CMD (ambiente em que a tese foi processada), com
os caminhos de `Parametros_LDO_LOA.xlsx`, da pasta raiz e do Tesseract fixados no topo de
`analise_ldo_loa.py` (bloco "ÁREA DE ALTERAÇÃO"). Para rodar em Linux ou Mac, edite esse bloco
(linhas ~23-28) trocando os caminhos `C:\...` pelos caminhos equivalentes no seu sistema (por
exemplo `/home/usuario/analiseldoloa/...`) e pelo caminho do binário `tesseract` do seu sistema
(`which tesseract`). Fora isso, o restante do script é multiplataforma.

## Erros comuns

- `No module named ...` → repita o passo 3 (`py -m pip install -r requirements.txt`).
- `Tesseract não encontrado` → confirme o passo 2 e o idioma "por" instalado.
- `Permission denied` ao gravar a saída → feche `Resultados_LDO_LOA.xlsx` no Excel antes de rodar
  de novo.
- Erro lendo parâmetros → não renomeie abas ou cabeçalhos de `Parametros_LDO_LOA.xlsx`; feche o
  Excel antes de rodar.
- Demora muito → normal com OCR em documentos grandes; para um teste rápido, na aba
  `Parametros_Gerais` da planilha de parâmetros, ajuste `USAR_OCR = Nao`.

## O que o programa faz (resumo)

Lê os parâmetros → extrai texto dos PDFs (com OCR quando necessário) → busca o vocabulário →
classifica em dois eixos (força × ancoragem) → pondera → deduplica → calcula os indicadores e o
Grau de Integração Orçamentária → grava o Excel.

## Reprodutibilidade

Vocabulário, pesos e rubrica idênticos aos Anexos A, B e C do `Texto-Tese-Modificado.docx`. Os
resultados publicados na tese (item 4.3) foram gerados com Python 3.12, Tesseract 5 (idioma
`por`) e as versões mínimas de bibliotecas listadas em `requirements.txt`. A amostragem do
`Harness_Validacao.py` usa semente fixa (`SEMENTE = 42`).
