# Análises da tese — Baixada Fluminense Histórica

Códigos das análises documentais desenvolvidas na tese de doutorado de **Almeida, Yasmin Viana (2026)**,
sobre a institucionalização da geoinformação na gestão pública municipal da Baixada Fluminense Histórica.

O repositório reúne as duas rotinas de análise automatizada descritas no Capítulo 3 da tese, além de
uma interface gráfica que permite executá-las sem edição do código-fonte.

| Rotina | Item da tese | O que faz |
|---|---|---|
| **Instrumentos urbanísticos** | 3.2 | Lê os PDFs de CTM, Planta Genérica de Valores, Plano Diretor, Plano de Mobilidade, Plano de Saneamento e Plano Diretor de Tecnologia dos 8 municípios; mede o Grau de Integração Municipal da geoinformação. |
| **LDO e LOA** | 3.3 | Lê os PDFs das Leis de Diretrizes Orçamentárias e Leis Orçamentárias Anuais, organizados em pastas por município; mede o Grau de Integração Orçamentária da geoinformação. |

Os **documentos analisados não estão neste repositório** — são documentos públicos, pesados, obtidos
junto às prefeituras e aos diários oficiais. A composição do corpo documental está descrita nos itens
3.2.1 e 3.3.1 da tese.

---

## Estrutura

```
app.py                      Interface gráfica (executa as duas rotinas)
requirements.txt            Bibliotecas Python necessárias
parametros/
  Parametros_LDO_LOA.xlsx   Vocabulário, pesos, temas e rubrica (Anexos A, B e C da tese)
scripts/
  instrumentos/             Rotina do item 3.2
  ldo_loa/                  Rotina do item 3.3

```

A planilha `Parametros_LDO_LOA.xlsx` é o núcleo metodológico: concentra o vocabulário de busca, as
forças técnicas, os temas e a rubrica de classificação, de modo que esses parâmetros possam ser
inspecionados e alterados sem tocar no código.

---

## Instalação

**1. Python 3.12 ou superior** — [python.org/downloads](https://www.python.org/downloads/).
No Windows, marque **"Add python.exe to PATH"** durante a instalação.

**2. Tesseract OCR** — necessário para ler os PDFs digitalizados como imagem.
Instalador para Windows: [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki).
Em *Additional language data*, marque **Portuguese** e **English**.

> O instalador do Tesseract **não** adiciona o programa ao PATH do Windows. É normal precisar
> informar o caminho do `tesseract.exe` na tela de parâmetros da interface — o padrão é
> `C:\Program Files\Tesseract-OCR\tesseract.exe`.

**3. Bibliotecas Python** — no terminal, dentro da pasta do repositório:

```bash
pip install -r requirements.txt
```

## Como executar

```bash
python app.py
```

A interface conduz a execução em quatro etapas: escolha da rotina, seleção da pasta de documentos,
conferência dos parâmetros (já preenchidos com os valores adotados na pesquisa) e execução com
registro em tela. Ao final, abre a pasta de resultados ou a planilha de evidências gerada.

As rotinas também podem ser executadas diretamente, sem a interface — ver
`documentacao/README_Instrumentos_original.md` e `documentacao/README_LDO_LOA_original.md`.

---

## Notas metodológicas

Registradas aqui por transparência, para quem pretenda reproduzir ou auditar os resultados.

**Idioma do OCR.** Na rotina das LDO/LOA o idioma do reconhecimento óptico é lido da planilha de
parâmetros (`IDIOMA_OCR`, definido como `por`). 

**Corpo documental.** A rotina dos instrumentos identifica município e tipo de instrumento a partir do
**nome de cada arquivo**, no formato `<Tipo> - <Município>.pdf` (por exemplo, `Plano Diretor - Japeri.pdf`).
Renomear os arquivos altera a classificação. A rotina das LDO/LOA, por sua vez, identifica os
municípios pela estrutura de subpastas, o que permite aplicá-la a outros recortes territoriais.

**Ajustes de portabilidade.** Os scripts foram originalmente desenvolvidos em ambiente Linux. Para
que executem também em Windows, foram explicitados a codificação de caracteres na leitura e escrita
dos arquivos intermediários, substituídos caminhos absolutos por variáveis configuráveis e
centralizada a configuração do mecanismo de OCR entre as etapas. Esses ajustes não alteram o
vocabulário, os pesos, a rubrica de classificação nem os critérios de identificação das ocorrências.
As versões anteriores dos arquivos permanecem acessíveis no histórico de commits deste repositório.

---

## Como citar

> ALMEIDA, Yasmin Viana. **A institucionalização da geoinformação e suas assimetrias territoriais na gestão pública municipal da Baixada Fluminense Histórica**. 2026. Tese (Doutorado em Geografia) —
> Universidade do Estado do Rio de Janeiro, Rio de Janeiro, 2026.

---

## Licença

Consulte a autora quanto às condições de uso e reaproveitamento do código.
