# -*- coding: utf-8 -*-
"""Reprocessamento pós-parecer:
- exclui documentos contaminados do corpus principal;
- aplica falso-positivo e cadastro-fiscal-fraco (antes carregados, não aplicados);
- deduplica boilerplate ENTRE documentos;
- recomputa categoria por função única (sem 'indícios periféricos');
- recalcula o Grau de densidade documental só com ocorrências APROVEITÁVEIS
  (não-periféricas e não-boilerplate) e apenas nos instrumentos VIGENTES.
Gera resultados_v2.json e ocorrencias_v2.json.
"""
import json, os, unicodedata, re
from collections import defaultdict, Counter

# Planilha de parametros (Parametros_LDO_LOA.xlsx), usada abaixo para reler a aba
# Grau_Rubrica. Por padrao, procura o arquivo na pasta atual.
PARAMETROS_XLSX = os.environ.get("PARAMETROS_XLSX", "Parametros_LDO_LOA.xlsx")

OCC=json.load(open("ocorrencias.json", encoding="utf-8"))
AF=json.load(open("analise_final.json", encoding="utf-8"))
AP=json.load(open("anexo_params.json", encoding="utf-8"))
meta={x["arquivo"]:x for x in AF["instrumentos"]}

EXCLUIR={"CTM - Japeri.pdf","PGV - Nilopolis - Noticia.pdf"}
FP=[unicodedata.normalize("NFKD",e).encode("ascii","ignore").decode().lower()
    for e in AP["apoio"].get("EXPRESSOES_FALSO_POSITIVO",[])]
FRACO=[unicodedata.normalize("NFKD",e).encode("ascii","ignore").decode().lower()
    for e in AP["apoio"].get("EXPRESSOES_CADASTRO_FISCAL_FRACO",[])]

MUNIS=["Belford Roxo","Duque de Caxias","Japeri","Mesquita","Nilopolis","Nova Iguaçu","Queimados","São João de Meriti"]

# ---------- 1) filtra corpus e falso-positivo ----------
occ=[]; n_excl_doc=0; n_fp=0
for o in OCC:
    if o["arquivo"] in EXCLUIR:
        n_excl_doc+=1; continue
    ctx=o["contexto"]
    if any(e in ctx for e in FP):
        n_fp+=1; continue
    occ.append(dict(o))

# ---------- 2) cadastro-fiscal-fraco (rebaixa peso) ----------
n_fraco=0
for o in occ:
    if o["termo"]=="cadastro fiscal" and any(e in o["contexto"] for e in FRACO):
        o["peso"]=round(o["peso"]*0.5,2); o["fraco"]=True; n_fraco+=1
    else:
        o["fraco"]=False

# ---------- 3) boilerplate entre documentos ----------
grp=defaultdict(set)
for o in occ: grp[(o["termo"],o["contexto"][:160])].add(o["arquivo"])
n_boiler=0
for o in occ:
    inter = len(grp[(o["termo"],o["contexto"][:160])])>1
    o["boilerplate"]=bool(inter)
    if inter: n_boiler+=1

# ---------- 4) recomputa por instrumento (função única) ----------
CAM_ORD={"DIRETO":3,"GENERICO":2,"PRATICA":1,"PERIFERICO":0}
def categoria(cams,nd):
    if "DIRETO" in cams: return "Explícita (forte)" if nd>=2 else "Explícita (fraca, 1 menção)"
    if "GENERICO" in cams: return "Implícita"
    if "PRATICA" in cams: return "Prática (sem terminologia)"
    return "Ausente (só termos periféricos ou nada)"

por_arq=defaultdict(list)
for o in occ: por_arq[o["arquivo"]].append(o)

res=[]
for arq,m in meta.items():
    if arq in EXCLUIR: continue
    dococc=por_arq.get(arq,[])
    evid=[o for o in dococc if not o["boilerplate"]]           # evidência: sem boilerplate
    aprov=[o for o in evid if o["camada"]!="PERIFERICO"]        # aproveitável: sem periférico
    cams_evid=set(o["camada"] for o in evid if o["camada"]!="PERIFERICO")
    nd=sum(1 for o in evid if o["camada"]=="DIRETO")
    cat=categoria(cams_evid,nd)
    tem_d="DIRETO" in cams_evid; tem_g="GENERICO" in cams_evid; tem_p="PRATICA" in cams_evid
    res.append({**{k:m[k] for k in ("arquivo","municipio","instrumento","ano_efetivo","vigencia_juridica",
                    "operante","situacao","revisao_status","pd_defasado","leis_federais","n_paginas")},
        "categoria":cat,"eixo_precisao":("Direta" if tem_d else "Genérica" if tem_g else "Nenhuma"),
        "eixo_pratica":("Com prática" if tem_p else "Sem prática"),
        "n_direto":nd,
        "n_ocorr_evid":len(evid),
        "n_aproveitavel":len(aprov),
        "soma_peso_aprov":round(sum(o["peso"] for o in aprov),1),
        "forca_max_aprov":max([o["forca"] for o in aprov],default=0),
        "temas_aprov":"; ".join(sorted(set(o["tema"] for o in aprov))),
        "n_boilerplate":sum(1 for o in dococc if o["boilerplate"]),
    })
res_by=defaultdict(list)
for r in res: res_by[r["municipio"]].append(r)

# ---------- 5) Grau de DENSIDADE documental (só operantes + aproveitáveis) ----------
import openpyxl
wbp=openpyxl.load_workbook(PARAMETROS_XLSX,data_only=True)
ws=wbp["Grau_Rubrica"];rows=list(ws.iter_rows(values_only=True))
hdr=[str(x).strip().lower() if x else "" for x in rows[1]]
RUB=[dict(zip(hdr,r)) for r in rows[2:] if r and r[0]]
def num(v):
    try:return float(v)
    except:return 0
def grau(soma,na,nanos,ntem,forte):
    for r in RUB:
        if str(r.get("nivel")).lower()=="ausente": continue
        ok=(soma>=num(r["min_soma_peso"]) and na>=num(r["min_ocorr_aproveitaveis"])
            and nanos>=num(r["min_anos_com_evidencia"]) and ntem>=num(r["min_temas_distintos"]))
        if str(r.get("exige_evidencia_forte","")).lower() in ("sim","s") and not forte: ok=False
        if ok: return r["nivel"]
    return "Ausente"

graus={}
for mu in MUNIS:
    rs=[r for r in res_by.get(mu,[]) if r["operante"]]
    soma=sum(r["soma_peso_aprov"] for r in rs)
    na=sum(r["n_aproveitavel"] for r in rs)
    anos=set(r["ano_efetivo"] for r in rs if r["ano_efetivo"] and r["n_aproveitavel"]>0)
    temas=set()
    for r in rs:
        if r["temas_aprov"]: temas.update(r["temas_aprov"].split("; "))
    forte=any(r["forca_max_aprov"]>=3 for r in rs)
    pgs=sum(r["n_paginas"] for r in rs)
    graus[mu]={"grau_densidade":grau(soma,na,len(anos),len([t for t in temas if t]),forte),
        "soma_aprov":round(soma,1),"n_aprov":na,"n_anos":len(anos),"n_temas":len([t for t in temas if t]),
        "forte":forte,"norm":round(1000*soma/pgs,1) if pgs else 0}

out={"excluidos":sorted(EXCLUIR),"instrumentos":res,"graus_densidade":graus,
     "contagens":{"occ_apos_filtro":len(occ),"fp_removidas":n_fp,"excl_doc":n_excl_doc,
                  "boilerplate":n_boiler,"cadastro_fiscal_fraco":n_fraco},
     "historico":[r["arquivo"] for r in res if not r["operante"]]}
json.dump(out,open("resultados_v2.json","w",encoding="utf-8"),ensure_ascii=False,indent=2)
json.dump(occ,open("ocorrencias_v2.json","w",encoding="utf-8"),ensure_ascii=False,indent=2)

# ---------- relatório ----------
print("== FILTRAGEM ==")
print(f"  ocorrências: {len(OCC)} -> {len(occ)} (removidas: {n_excl_doc} de doc excluído + {n_fp} falso-positivo)")
print(f"  boilerplate entre docs: {n_boiler} | cadastro fiscal fraco: {n_fraco}")
print("\n== CATEGORIAS (novas, corpus limpo) ==")
print("  ",dict(Counter(r["categoria"].split(" (")[0] for r in res)))
print("\n== JAPERI e NILÓPOLIS (afetados) ==")
for r in res:
    if r["municipio"] in ("Japeri","Nilopolis"):
        print(f"   {r['instrumento']:16} {r['categoria']:32} aprov={r['n_aproveitavel']}")
print("\n== GRAU DE DENSIDADE (antes: quase todos Consolidado) ==")
for mu in MUNIS:
    g=graus[mu]
    print(f"   {mu:18} {g['grau_densidade']:12} soma_aprov={g['soma_aprov']:6} n_aprov={g['n_aprov']:3} temas={g['n_temas']} forte={g['forte']}")
