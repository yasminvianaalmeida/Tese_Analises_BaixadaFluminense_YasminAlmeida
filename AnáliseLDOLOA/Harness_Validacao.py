# -*- coding: utf-8 -*-
# ==================================================================================
# HARNESS DE VALIDACAO DA CLASSIFICACAO AUTOMATICA  ---  LDO/LOA v2
# ==================================================================================
# Mede CONFIABILIDADE e VALIDADE da classificacao automatica.
#   MODO "amostrar": sorteia amostra estratificada e gera planilha para codificacao manual.
#   MODO "avaliar":  le a planilha preenchida e calcula kappa de Cohen (humano x humano),
#                    precisao/revocacao/F1 (maquina x gabarito humano) e matriz de confusao.
# Dependencias: pandas, openpyxl, scikit-learn
# USO:
#   py Harness_Validacao.py amostrar CONSOLIDADO_v2_Analise.xlsx amostra_para_codificar.xlsx
#   py Harness_Validacao.py avaliar  amostra_para_codificar.xlsx relatorio_validacao.xlsx
# ==================================================================================
import sys, unicodedata
import pandas as pd

TAMANHO_AMOSTRA = 200          # >>> DECISAO: nº de ocorrencias para codificacao manual
SEMENTE = 42
COLUNA_ESTRATO = "classe"
ROTULOS_VALIDOS = [
    "ocorrencia territorial relevante",
    "ocorrencia validada (tecnica)",
    "ocorrencia contextual",
    "ocorrencia administrativa generica / falso positivo",
]

def amostrar(arquivo_consolidado, arquivo_saida):
    xls = pd.ExcelFile(arquivo_consolidado)
    aba = "04_ocorrencias" if "04_ocorrencias" in xls.sheet_names else ("Ocorrencias_Relevantes" if "Ocorrencias_Relevantes" in xls.sheet_names else xls.sheet_names[0])
    df = pd.read_excel(arquivo_consolidado, sheet_name=aba)
    grupos = []
    for _, g in df.groupby(COLUNA_ESTRATO, dropna=False):
        n = max(1, round(TAMANHO_AMOSTRA * len(g) / len(df)))
        grupos.append(g.sample(min(n, len(g)), random_state=SEMENTE))
    amostra = pd.concat(grupos).sample(frac=1, random_state=SEMENTE).reset_index(drop=True)
    cols = ["municipio","ano","tipo_documento","pagina","termo_chave","termo_encontrado",
            "trecho","forca_tecnica","ancoragem_territorial","peso_analitico","classe"]
    cols = [c for c in cols if c in amostra.columns]
    amostra = amostra[cols].rename(columns={"classe":"classe_maquina"})
    amostra["rotulo_humano_1"] = ""
    amostra["rotulo_humano_2"] = ""
    amostra["observacao_humano"] = ""
    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as w:
        amostra.to_excel(w, sheet_name="codificar", index=False)
        pd.DataFrame({"rotulos_validos": ROTULOS_VALIDOS}).to_excel(w, sheet_name="instrucoes", index=False)
    print("Amostra com", len(amostra), "ocorrencias salva em:", arquivo_saida)

def _canon(s):
    """Normaliza rotulo: sem acento, minusculo, espacos colapsados. Mapeia para os 4 padrao."""
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()
    s = " ".join(s.split())
    # sinonimos/abreviacoes comuns -> rotulo canonico
    if s in ("", "nan", "none"): return ""
    if "territorial relevante" in s: return "territorial relevante"
    if "validada" in s: return "validada (tecnica)"
    if "contextual" in s: return "contextual"
    if "administrativa" in s or "falso positivo" in s or "generica" in s: return "administrativa/falso positivo"
    return s  # rotulo nao reconhecido (sera sinalizado)

def avaliar(arquivo_preenchido, arquivo_relatorio):
    try:
        from sklearn.metrics import cohen_kappa_score, precision_recall_fscore_support, confusion_matrix
    except ImportError:
        sys.exit("Instale o scikit-learn: py -m pip install scikit-learn")
    df = pd.read_excel(arquivo_preenchido, sheet_name="codificar")
    for c in ["rotulo_humano_1", "rotulo_humano_2", "observacao_humano", "classe_maquina"]:
        if c not in df.columns: df[c] = ""
    # normaliza tudo para rotulo canonico (resolve acento/caixa/sinonimo)
    df["h1"]  = df["rotulo_humano_1"].map(_canon)
    df["h2"]  = df["rotulo_humano_2"].map(_canon)
    df["maq"] = df["classe_maquina"].map(_canon)
    df = df[df["h1"] != ""]
    if df.empty:
        sys.exit("Nenhuma linha codificada manualmente foi encontrada (coluna rotulo_humano_1).")
    validos = {"territorial relevante","validada (tecnica)","contextual","administrativa/falso positivo"}
    nrec = sorted(set(df["h1"]) - validos)
    if nrec:
        print("AVISO: rotulos humanos nao reconhecidos (verifique grafia):", nrec)
    h1  = [str(x) for x in df["h1"].tolist()]
    maq = [str(x) for x in df["maq"].tolist()]
    rel = {}

    # CONFIABILIDADE: linhas com os DOIS humanos
    sub = df[df["h2"] != ""]
    if len(sub) >= 2:
        a = [str(x) for x in sub["h1"].tolist()]
        b = [str(x) for x in sub["h2"].tolist()]
        try:
            kappa = float(cohen_kappa_score(a, b))
        except Exception as e:
            kappa = float("nan"); print("kappa falhou:", e)
        concord = sum(1 for x, y in zip(a, b) if x == y) / len(a)
        print("[CONFIABILIDADE] Kappa de Cohen:", round(kappa, 3), "| concordancia:", round(concord, 3), "| n =", len(sub))
        rel["confiabilidade"] = pd.DataFrame({"metrica": ["kappa_cohen", "concordancia", "n"],
                                              "valor": [round(kappa, 3), round(concord, 3), len(sub)]})
    else:
        print("[CONFIABILIDADE] pulada: menos de 2 linhas com o 2o avaliador (coluna rotulo_humano_2).")

    # VALIDADE: maquina x gabarito humano
    rotulos = sorted(set(h1) | set(maq))
    p, r, f, _ = precision_recall_fscore_support(h1, maq, labels=rotulos, zero_division=0)
    val = pd.DataFrame({"classe": rotulos, "precisao": p.round(3), "revocacao": r.round(3), "f1": f.round(3)})
    pm, rm, fm, _ = precision_recall_fscore_support(h1, maq, average="macro", zero_division=0)
    acc = sum(1 for x, y in zip(h1, maq) if x == y) / len(h1)
    print("[VALIDADE] Acuracia:", round(acc, 3), "| F1 macro:", round(float(fm), 3), "| n =", len(h1))
    print(val.to_string(index=False))
    rel["validade_por_classe"] = val
    rel["validade_global"] = pd.DataFrame({"metrica": ["acuracia", "precisao_macro", "revocacao_macro", "f1_macro", "n"],
                                           "valor": [round(acc, 3), round(float(pm), 3), round(float(rm), 3), round(float(fm), 3), len(h1)]})
    rel["matriz_confusao"] = pd.DataFrame(confusion_matrix(h1, maq, labels=rotulos), index=rotulos, columns=rotulos)
    with pd.ExcelWriter(arquivo_relatorio, engine="openpyxl") as w:
        for nome, tab in rel.items():
            tab.to_excel(w, sheet_name=nome[:31], index=(nome == "matriz_confusao"))
    print("Relatorio salvo em:", arquivo_relatorio)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit("Uso: amostrar|avaliar  <entrada>  <saida>")
    modo = sys.argv[1].lower()
    if modo == "amostrar": amostrar(sys.argv[2], sys.argv[3])
    elif modo == "avaliar": avaliar(sys.argv[2], sys.argv[3])
    else: sys.exit("Modo invalido.")
