# -*- coding: utf-8 -*-
"""Pacote de robustez + corte >=2015 + modalidade. Gera analise_final.json."""
import fitz, re, glob, os, json, unicodedata
from collections import Counter, defaultdict

# ================== AREA DE ALTERACAO (somente o caminho) =========================
# Aponte para a pasta com os PDFs dos instrumentos urbanisticos (a mesma pasta passada
# com --pasta ao rodar analise_instrumentos.py). Ajuste este caminho para o seu ambiente
# antes de rodar (por padrao, aponta para uma subpasta "Instrumentos" ao lado deste script).
PASTA = os.environ.get("PASTA_INSTRUMENTOS", os.path.join(os.path.dirname(os.path.abspath(__file__)), "Instrumentos"))
# ====================================================================================
R=json.load(open("resultados.json")); OCC=json.load(open("ocorrencias.json"))
inst={x["arquivo"]:x for x in R["instrumentos"]}
CORTE=2015
cache={"PGV - Nilopolis.pdf":"cache_pgv_nilo.json","Plano Diretor - Nilopolis.pdf":"cache_pd_nilo.json"}
MES=r"(?:janeiro|fevereiro|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)"

def norm(t):
    t=unicodedata.normalize("NFKD",t); t="".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+"," ",t.lower())
def first_text(path,fn,npg=4):
    if fn in cache:
        c=json.load(open(cache[fn]))["paginas"]
        return " ".join(c[str(k)] for k in sorted(int(x) for x in c)[:npg])
    d=fitz.open(path); t=" ".join(d[i].get_text() for i in range(min(npg,d.page_count))); d.close()
    return norm(t)
def meta_year(path):
    try:
        d=fitz.open(path); m=d.metadata or {}; d.close()
        for k in ("creationDate","modDate"):
            mm=re.search(r"(\d{4})", m.get(k) or "")
            if mm:
                y=int(mm.group(1))
                if 1990<=y<=2026: return y
    except Exception: pass
    return None

# ---------- 1) DATACAO ----------
datas={}
for path in sorted(glob.glob(os.path.join(PASTA,"*.pdf"))):
    fn=os.path.basename(path); ft=first_text(path,fn)
    inst_y=[int(a) for a in re.findall(rf"(?:lei(?: complementar)?|decreto)[^.]{{0,80}}?de\s+\d{{1,2}}\s+de\s+{MES}\s+de\s+(\d{{4}})",ft) if 1990<=int(a)<=2026]
    inst_y=max(inst_y) if inst_y else None
    capa=[int(a) for a in re.findall(r"\b(19\d{2}|20[0-2]\d)\b",ft) if 1990<=int(a)<=2026]
    capa_max=max(capa) if capa else None
    meta=meta_year(path)
    # decisao de ano e confianca
    if inst_y:
        ano=inst_y; conf="alta"; ev=f"data da norma instituidora ({inst_y})"
    elif capa_max and meta and abs(capa_max-meta)<=1:
        ano=max(capa_max,meta); conf="alta"; ev=f"capa {capa_max} + metadados {meta}"
    elif capa_max:
        ano=capa_max; conf="media"; ev=f"ano mais recente no preambulo ({capa_max})"
    elif meta:
        ano=meta; conf="media"; ev=f"metadados do PDF ({meta})"
    else:
        ano=None; conf="baixa"; ev="sem sinal claro"
    # flag de conflito (instituidora antiga x referencias recentes)
    if inst_y and capa_max and capa_max-inst_y>=5:
        conf="CONFIRMAR"; ev+=f"; ATENCAO: norma de {inst_y} mas referencias ate {capa_max}"
    datas[fn]={"ano":ano,"conf":conf,"evidencia":ev,"inst_y":inst_y,"capa_max":capa_max,"meta":meta}

# ---------- 2) MODALIDADE (forca 3) ----------
RX_ESTR=re.compile(r"\bfica(?:m)? (?:criad[oa]s?|institu[ido]+|estabelecid[oa]s?|implantad[oa]s?)\b|\binstitui-se\b|\bcompete\b.{0,40}\bmanter\b|\bmantera\b|\bimplantara\b|\bo municipio (?:criara|implantara|mantera)\b")
RX_ASP=re.compile(r"\bser[ao]+ (?:criad[oa]s?|implantad[oa]s?|elaborad[oa]s?)\b|\bpoder[ao]+\b|\bdever[ao]+ ser\b|\bobjetivo de\b|\bdiretriz\b|\bpropoe\b|\bvisa\b|\bbuscar[ao]+\b|\bincentivar\b|\bpretende\b|\brecomenda")
RX_DESC=re.compile(r"\bpor meio de\b|\batraves de\b|\butilizou-se\b|\bfoi(?:ram)? (?:realizad|utilizad|elaborad)|\bsoftware\b|\bferramenta(?:s)? de\b|\bcom base em\b")
def modalidade(ctx):
    if RX_ESTR.search(ctx): return "Estruturante (institui/cria)"
    if RX_ASP.search(ctx): return "Aspiracional (prevê/pretende)"
    if RX_DESC.search(ctx): return "Descritiva (uso/método)"
    return "Mencao simples"

# ---------- 3) BOILERPLATE federal ----------
RX_BOIL=re.compile(r"incluido pela lei|redacao dada pela|nos termos da lei|na forma da lei|conforme (?:a )?lei n|(?:art(?:igo)?\.?\s*\d+)")
# checa se citacoes federais aparecem em contexto de transcricao legal
def leis_boilerplate(texto_norm, leis):
    res={}
    for lei in leis:
        # encontra a sigla/numero e olha o redor
        pat={"Estatuto da Metropole (Lei 13.089/2015)":r"13\.?089",
             "Lei 13.683/2018 (altera Estatuto Metropole)":r"13\.?683",
             "Estatuto da Cidade (Lei 10.257/2001)":r"10\.?257"}.get(lei)
        if not pat: res[lei]="—"; continue
        boil=0; total=0
        for m in re.finditer(pat,texto_norm):
            total+=1; ctx=texto_norm[max(0,m.start()-120):m.end()+40]
            if RX_BOIL.search(ctx): boil+=1
        if total==0: res[lei]="—"
        elif boil>=total*0.6: res[lei]="boilerplate (citação herdada)"
        else: res[lei]="deliberada"
    return res

# texto completo por arquivo (para boilerplate) - reusa ocorrencias? precisamos do texto.
# Aproximacao: usamos contexto das ocorrencias ja extraidas + primeiras paginas.
texto_por_arq=defaultdict(str)
for o in OCC: texto_por_arq[o["arquivo"]]+=" "+o["contexto"]

# ---------- 4) reclassificacao com PISO e DOIS EIXOS ----------
direto_count=defaultdict(int); cam_por_arq=defaultdict(set); occ_por_arq=defaultdict(list)
for o in OCC:
    occ_por_arq[o["arquivo"]].append(o)
    if o["camada"]!="PERIFERICO": cam_por_arq[o["arquivo"]].add(o["camada"])
    if o["camada"]=="DIRETO": direto_count[o["arquivo"]]+=1

def reclassifica(fn):
    cams=cam_por_arq[fn]; nd=direto_count[fn]
    tem_d="DIRETO" in cams; tem_g="GENERICO" in cams; tem_p="PRATICA" in cams
    # eixo precisao
    precisao="Direta" if tem_d else ("Genérica" if tem_g else "Nenhuma")
    # eixo pratica
    pratica="Com prática espacial" if tem_p else "Sem prática"
    # categoria com piso
    if tem_d:
        cat="Explícita (forte)" if nd>=2 else "Explícita (fraca, 1 menção)"
    elif tem_g:
        cat="Implícita"
    elif tem_p:
        cat="Prática (sem terminologia)"
    else:
        cat="Ausente (só termos periféricos ou nada)"
    return cat,precisao,pratica,tem_d,tem_g,tem_p,nd

# ---------- monta resultado final ----------
LEIS=["Estatuto da Cidade (Lei 10.257/2001)","INDE - Decreto 6.666/2008","Decreto 12.402/2025 (atualiza INDE)",
 "Politica Nacional de Desenvolvimento Urbano (PNDU)","Estatuto da Metropole (Lei 13.089/2015)",
 "Lei 13.683/2018 (altera Estatuto Metropole)","Portaria 3.242/2022 (CTM)","IN RFB 2.030/2021 (CIB)","Decreto 11.208/2022 (SINTER)"]
ANO_LEI={"Estatuto da Cidade (Lei 10.257/2001)":2001,"INDE - Decreto 6.666/2008":2008,
 "Decreto 12.402/2025 (atualiza INDE)":2025,"Politica Nacional de Desenvolvimento Urbano (PNDU)":2001,
 "Estatuto da Metropole (Lei 13.089/2015)":2015,"Lei 13.683/2018 (altera Estatuto Metropole)":2018,
 "Portaria 3.242/2022 (CTM)":2022,"IN RFB 2.030/2021 (CIB)":2021,"Decreto 11.208/2022 (SINTER)":2022}

# ---------- SITUACAO / VIGENCIA (em vez de excluir <2015) ----------
CUR=2026
OVERRIDE={
 "PGV - Belford Roxo.pdf":{"ano_base":2005,"ano_atualizacao":2023,
   "obs":"Lei-base de 2005, atualizada até 2023; vários artigos alterados/revogados por leis complementares."},
 "PGV - Nilopolis.pdf":{"ano_base":2020,"ano_atualizacao":2020,"obs":"Correção manual: instrumento de 2020."},
}
HIST_SUBST={"Plano Diretor - Queimados-Antigo.pdf":"Plano Diretor - Queimados - Atual",
            "Plano Diretor - São João de Meriti Antigo.pdf":"Plano Diretor - São João de Meriti"}
def situacao_doc(fn,x):
    d=datas[fn]; ov=OVERRIDE.get(fn,{})
    base = ov.get("ano_base", d["inst_y"] or d["capa_max"])
    atual = ov.get("ano_atualizacao", d["capa_max"])  # ultimo ano citado no texto (proxy)
    cand=[y for y in (base,atual) if y]
    efet = min(max(cand),CUR) if cand else None
    obs = ov.get("obs","")
    if efet is None and d["meta"]:
        efet=min(d["meta"],CUR); base=base or efet; atual=atual or efet
        obs=(obs+f" Ano estimado por metadados do PDF ({d['meta']}) — confirmar.").strip()
    if "noticia" in fn.lower(): obs=(obs+" Fonte secundária (notícia).").strip()
    substituido = fn in HIST_SUBST
    # EIXO 1 - vigencia juridica (a norma esta em vigor?)
    if substituido:
        vigencia="Substituído (revogado por versão posterior)"; obs=(obs+f" Substituído por: {HIST_SUBST[fn]}.").strip()
    else:
        vigencia="Vigente"
    is_pd = x["instrumento"]=="Plano Diretor"
    defas = (is_pd and efet and (CUR-efet)>10 and not substituido)
    # EIXO 2 - revisao/atualidade (default editavel; leis complementares != revisao)
    base_antiga = bool((base and base<CORTE))
    atualizado_compl = base_antiga and atual and atual>=CORTE
    if substituido:
        rev="—"
    elif is_pd and defas:
        rev="Vigente, SEM revisão completa — defasado (art. 40 §3º); verificar leis complementares"
    elif atualizado_compl:
        rev="Vigente — mantido por leis complementares (sem revisão completa)"
    elif efet and (CUR-efet)<=10:
        rev="Norma recente"
    else:
        rev="Base antiga — verificar atualização"
    fonte_proxy = (fn not in OVERRIDE)
    return base,atual,efet,vigencia,rev,defas,obs,fonte_proxy

final=[]
for fn,x in inst.items():
    d=datas[fn]; cat,prec,prat,td,tg,tp,nd=reclassifica(fn)
    base,atual,efet,vigencia,rev,defas,obs,fproxy=situacao_doc(fn,x)
    operante = not vigencia.startswith("Substituído")
    boil=leis_boilerplate(texto_por_arq[fn], [l for l in LEIS if l in x["leis_federais"]])
    leis_citadas=[l for l in LEIS if l in x["leis_federais"]]
    leis_anacronicas=[l for l in leis_citadas if efet and ANO_LEI[l]>efet]
    final.append({**x,
        "ano_base":base,"ano_atualizacao":atual,"ano_efetivo":efet,"ano_atual_proxy":fproxy,
        "ano_conf":d["conf"],"ano_evidencia":d["evidencia"],
        "vigencia_juridica":vigencia,"revisao_status":rev,"operante":operante,
        "pd_defasado":bool(defas),"observacao":obs,
        "situacao":vigencia,"vigente":operante,"ano_documento":efet,  # compat
        "categoria_v2":cat,"eixo_precisao":prec,"eixo_pratica":prat,"n_ocorr_diretas":nd,
        "leis_boilerplate":"; ".join(f"{k.split('(')[0].strip()}={v}" for k,v in boil.items() if v not in("—",)) or "—",
        "leis_citadas_aplic":len([l for l in leis_citadas if l not in leis_anacronicas]),
    })
efet_de={x["arquivo"]:x["ano_efetivo"] for x in final}
sit_de={x["arquivo"]:x["vigencia_juridica"] for x in final}
vig_de={x["arquivo"]:x["operante"] for x in final}

# ---------- modalidade das forca-3 (todas, sem exclusao) ----------
mod=[]
for o in OCC:
    if o["forca"]==3:
        o2=dict(o); o2["ano_efetivo"]=efet_de.get(o["arquivo"])
        o2["situacao"]=sit_de.get(o["arquivo"]); o2["vigente"]=vig_de.get(o["arquivo"])
        o2["modalidade_auto"]=modalidade(o["contexto"]); o2["modalidade_validada"]=""
        mod.append(o2)

# ---------- GRAU bruto vs normalizado (corpus final) ----------
def num(v):
    try:return float(v)
    except:return 0
RUBRICA=R.get("rubrica") or []
# recarrega rubrica do excel de parametros
import openpyxl
wbp=openpyxl.load_workbook("/sessions/clever-beautiful-fermat/mnt/RevisaoMetodologia_Instrumentos Urbanísticos/Parametros_LDO_LOA.xlsx",data_only=True)
ws=wbp["Grau_Rubrica"]; rows=list(ws.iter_rows(values_only=True))
hdr=[str(c).strip().lower() if c else "" for c in rows[1]]
RUB=[dict(zip(hdr,r)) for r in rows[2:] if r and r[0]]
def grau(soma,na,nanos,ntem,forte):
    for r in RUB:
        if str(r.get("nivel")).lower()=="ausente": continue
        ok=(soma>=num(r["min_soma_peso"]) and na>=num(r["min_ocorr_aproveitaveis"])
            and nanos>=num(r["min_anos_com_evidencia"]) and ntem>=num(r["min_temas_distintos"]))
        if str(r.get("exige_evidencia_forte","")).lower() in("sim","s") and not forte: ok=False
        if ok: return r["nivel"]
    return "Ausente"

MUNIS=["Belford Roxo","Duque de Caxias","Japeri","Mesquita","Nilopolis","Nova Iguaçu","Queimados","São João de Meriti"]
def calc_graus(corpus):
    g={}
    pm=defaultdict(list)
    for x in corpus: pm[x["municipio"]].append(x)
    for m in MUNIS:
        rs=pm.get(m,[])
        if not rs: g[m]={"grau":"Ausente (sem doc no recorte)","soma":0,"pgs":0,"norm":0,"na":0}; continue
        soma=sum(r["soma_peso"] for r in rs); pgs=sum(r["n_paginas"] for r in rs)
        na=sum(r["n_ocorrencias"] for r in rs)
        anos=set(r["ano_efetivo"] for r in rs if r["ano_efetivo"])
        temas=set(); [temas.update(r["temas"].split("; ")) for r in rs if r["temas"]]
        forte=any(r["forca_max"]>=3 for r in rs)
        g[m]={"grau":grau(soma,na,len(anos),len([t for t in temas if t]),forte),
              "soma":round(soma,1),"pgs":pgs,"norm":round(1000*soma/pgs,1) if pgs else 0,
              "na":na,"nanos":len(anos),"ntemas":len([t for t in temas if t]),"forte":forte}
    return g

g_todos=calc_graus(final)
g_vig=calc_graus([x for x in final if x["vigente"]])

pd_status={}
for m in MUNIS:
    pds=[x for x in final if x["municipio"]==m and x["instrumento"]=="Plano Diretor" and x["operante"]]
    if not pds: pd_status[m]={"ano":None,"defasado":None,"idade":None,"categoria":None,"revisao":None}; continue
    op=max(pds,key=lambda a:a["ano_efetivo"] or 0)
    idade=(CUR-op["ano_efetivo"]) if op["ano_efetivo"] else None
    pd_status[m]={"ano":op["ano_efetivo"],"defasado":bool(idade and idade>10),"idade":idade,
                  "categoria":op["categoria_v2"],"revisao":op["revisao_status"]}

out={"corte_referencia":CORTE,"instrumentos":final,"forca3_modalidade":mod,
     "graus_todos":g_todos,"graus_vigentes":g_vig,"pd_status":pd_status,
     "n_total":len(final),
     "n_operantes":sum(1 for x in final if x["operante"]),
     "n_substituidos":sum(1 for x in final if not x["operante"]),
     "n_vigentes":sum(1 for x in final if x["operante"]),
     "n_historicos":sum(1 for x in final if not x["operante"]),
     "n_pd_defasados":sum(1 for m in pd_status if pd_status[m]["defasado"])}
json.dump(out,open("analise_final.json","w"),ensure_ascii=False,indent=2)

print("Total:",out["n_total"],"| Vigentes:",out["n_vigentes"],"| Historicos:",out["n_historicos"],"| PD defasados:",out["n_pd_defasados"])
print("\n=== SITUACAO POR DOCUMENTO ===")
for x in sorted(final,key=lambda a:(a["municipio"],a["instrumento"])):
    flag="[PD DEFASADO]" if x["pd_defasado"] else ""
    print("  efet=%-5s %-12s %-55s %-13s %s"%(x["ano_efetivo"],x["vigencia_juridica"][:11],x["revisao_status"][:53],flag,x["arquivo"][:34]))
print("\n=== PLANO DIRETOR: regra decenal ===")
for m in MUNIS:
    s=pd_status[m]
    print("  %-18s PD %-6s idade=%-4s %-22s cat=%s"%(m,s["ano"],s["idade"],"DEFASADO (>10)" if s["defasado"] else ("ok" if s["ano"] else "sem PD"),s.get("categoria")))
print("\n=== GRAU: todos vs operantes | norm ===")
for m in MUNIS:
    print("  %-18s %-13s %-13s %s"%(m,g_todos[m]["grau"][:12],g_vig[m]["grau"][:12],g_todos[m]["norm"]))
print("\nModalidade:",dict(Counter(o["modalidade_auto"] for o in mod)))
