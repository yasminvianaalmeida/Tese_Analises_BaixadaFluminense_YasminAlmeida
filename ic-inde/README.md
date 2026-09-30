# IC-INDE — planilha de replicação

Planilha que calcula o **Índice de Consolidação da Infraestrutura de Dados Espaciais (IC-INDE)** das
entidades cadastradas na INDE, descrito no item 3.1.1.3 da tese e documentado no Apêndice F.
Permite auditar os resultados do item 4.1 e recalcular o índice com outros pesos, outra meia-vida,
outra data de referência ou dados atualizados.

**Arquivo:** `IC-INDE_Planilha_Replicacao.xlsx` — baixe o arquivo e abra no Microsoft Excel ou importe no
Google Planilhas. O GitHub apenas armazena o arquivo e não recalcula fórmulas.

**Fonte dos dados:** Diretório Brasileiro de Dados Geoespaciais (DBDG) / INDE, extração de 25/06/2026.

## Abas

| Aba | Conteúdo | Editar? |
|---|---|---|
| `Base_bruta` | 71 registros institucionais, como extraídos do DBDG, sem alteração | Não |
| `Base_Original` | 69 registros após a consolidação dos dois pares homônimos (MB/DHN e MP) | Não |
| `Base_Tratada` | Os 69 registros com identificador, esfera, tipo de entidade e região | Sim, para atualizar os dados |
| `Parâmetros` | Pesos, meia-vida, data de referência e limiares das faixas | Sim |
| `Cálculo_IC-INDE` | Dimensões P, A e V, índice, faixa, ranking e cenário de pesos iguais, por registro | Não (fórmulas) |
| `Agregações` | Médias e distribuição por esfera, tipo de entidade e faixa | Não (fórmulas) |

## Parâmetros (aba `Parâmetros`)

| Parâmetro | Valor adotado | Célula |
|---|---|---|
| Peso da Publicização (wP) | 0,40 | B3 |
| Peso da Atualidade (wA) | 0,35 | B4 |
| Peso do Acervo (wV) | 0,25 | B5 |
| Meia-vida (τ) | 730 dias | B6 |
| Data de referência | 25/06/2026 | B7 |
| Limiares das faixas A, B, C e D | 80, 60, 40 e 20 | B12 a B15 |

A célula B10 indica `OK` quando os três pesos somam 1,00 e `AJUSTAR PESOS` em caso contrário.

## Fórmulas

- **P** (Publicização) = Publicados ÷ Total
- **A** (Atualidade) = mín[1; 0,5^(máx(0; Δt) ÷ τ)], com Δt = data de referência − última atualização (em dias)
- **V** (Acervo) = ln(1 + Total) ÷ ln(1 + M\*), com M\* = maior acervo observado
- **IC-INDE** = 100 × (wP·P + wA·A + wV·V)
- **Faixas:** A — Plena (≥ 80); B — Alta (60–79); C — Média (40–59); D — Baixa (20–39); E — Incipiente (< 20)

## Como atualizar

1. Na aba `Parâmetros`, altere pesos, meia-vida, data de referência ou limiares.
2. Para novos dados, substitua Total, Publicados, Pendentes e Última atualização na aba `Base_Tratada`,
   mantendo a ordem e a estrutura das colunas.
3. As abas `Cálculo_IC-INDE` e `Agregações` recalculam sozinhas.

## Limites

- As fórmulas cobrem os 69 registros atuais (linhas 2 a 70). Incluir novas instituições exige estender as
  fórmulas de `Cálculo_IC-INDE` e os intervalos de ranking e de M\*.
- `Base_Tratada` contém valores, não fórmulas derivadas de `Base_bruta`; a consolidação não é refeita
  automaticamente.
- O índice mede apenas o que as entidades declaram e publicam na INDE. Não avalia qualidade temática,
  metadados, interoperabilidade, geosserviços nem uso efetivo na decisão pública.
