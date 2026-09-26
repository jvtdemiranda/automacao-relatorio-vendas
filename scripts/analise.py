"""
Cálculos do relatório: consolida as lojas, compara com a meta e com o mês
anterior, e escreve os destaques em texto. Tudo sai num dicionário
("resumo") que alimenta o PDF, o Excel e a página — uma fonte só pros três.
"""

import csv
import math
import os
from dataclasses import asdict

from leitura import CATEGORIAS, LOJAS, OUTROS, valor_texto

NOMES_LOJAS = [loja for loja, _, _ in LOJAS]
MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
            "setembro", "outubro", "novembro", "dezembro"]


def ler_metas(caminho: str) -> dict:
    """
    {('2026-08', 'Centro'): 47000.0, ...}. Sem o arquivo, ou com uma meta
    ilegível, o relatório sai do mesmo jeito — só sem a comparação com meta.
    """
    if not os.path.exists(caminho):
        print(f"Aviso: {os.path.basename(caminho)} não encontrado — relatório sem comparação com meta")
        return {}
    metas = {}
    with open(caminho, newline="", encoding="utf-8-sig") as f:
        for l in csv.DictReader(f, delimiter=";"):
            meta = valor_texto(l.get("meta") or "")
            if math.isnan(meta) or not l.get("mes") or not l.get("loja"):
                print(f"Aviso: linha de meta ignorada: {l}")
                continue
            metas[(l["mes"].strip(), l["loja"].strip())] = meta
    return metas


def variacao(atual, anterior):
    if not anterior:
        return None
    return round((atual - anterior) / anterior, 4)


def numeros_da_loja(df):
    vendas = df[df["valor"] > 0]
    bruto = float(vendas["valor"].sum())
    devolucoes = float(-df.loc[df["valor"] < 0, "valor"].sum())
    # (loja, venda): duas lojas podem usar o mesmo número de cupom
    n_vendas = len(vendas[["loja", "venda"]].drop_duplicates())
    return {
        "bruto": round(bruto, 2),
        "devolucoes": round(devolucoes, 2),
        "liquido": round(bruto - devolucoes, 2),
        "vendas": n_vendas,
        # ticket médio sobre as vendas (devolução não é uma venda nova)
        "ticket_medio": round(bruto / n_vendas, 2) if n_vendas else 0.0,
    }


def resumir(dados, registros, metas, ano, mes, dados_anterior=None):
    chave_mes = f"{ano}-{mes:02d}"
    ano_ant, mes_ant = (ano, mes - 1) if mes > 1 else (ano - 1, 12)
    enviadas = {r.loja for r in registros if r.arquivo and r.linhas_validas}

    lojas = []
    for loja in NOMES_LOJAS:
        n = numeros_da_loja(dados[dados["loja"] == loja])
        ant = None
        if dados_anterior is not None:
            ant = numeros_da_loja(dados_anterior[dados_anterior["loja"] == loja])["liquido"] or None
        meta = metas.get((chave_mes, loja))
        lojas.append({
            "loja": loja, **n, "enviou": loja in enviadas,
            "meta": meta,
            "pct_meta": round(n["liquido"] / meta, 4) if meta else None,
            "liquido_anterior": ant,
            "variacao": variacao(n["liquido"], ant),
        })

    total = numeros_da_loja(dados)
    meta_total = sum(l["meta"] for l in lojas if l["meta"])
    ant_total = sum(l["liquido_anterior"] or 0 for l in lojas) or None
    total.update({
        "meta": meta_total or None,
        "pct_meta": round(total["liquido"] / meta_total, 4) if meta_total else None,
        "liquido_anterior": round(ant_total, 2) if ant_total else None,
        "variacao": variacao(total["liquido"], ant_total),
    })

    por_categoria = dados.groupby("categoria")["valor"].sum()
    categorias = []
    for cat in CATEGORIAS + [OUTROS]:
        valor = round(float(por_categoria.get(cat, 0.0)), 2)
        if valor or cat != OUTROS:
            categorias.append({"categoria": cat, "liquido": valor,
                               "participacao": round(valor / total["liquido"], 4) if total["liquido"] else 0})
    categorias.sort(key=lambda c: -c["liquido"])

    diario = dados.groupby("data")["valor"].sum().sort_index()
    diario = [{"data": d.isoformat(), "liquido": round(float(v), 2)} for d, v in diario.items()]

    resumo = {
        "mes": chave_mes,
        "mes_extenso": f"{MESES_PT[mes - 1]}/{ano}",
        "mes_anterior_extenso": f"{MESES_PT[mes_ant - 1]}/{ano_ant}",
        "total": total,
        "lojas": lojas,
        "categorias": categorias,
        "diario": diario,
        "qualidade": [dict(asdict(r), problemas=problemas(r)) for r in registros],
    }
    resumo["destaques"] = destaques(resumo)
    return resumo


def problemas(reg):
    """Tudo o que foi corrigido ou descartado numa loja, em frases prontas."""
    itens = [f"{motivo}: {n}" for motivo, n in reg.descartes.items()]
    if reg.devolucoes:
        itens.append(f"devoluções abatidas: {reg.devolucoes} ({brl(reg.valor_devolucoes)})")
    return itens + reg.avisos


def brl(v):
    # espaço não-separável: "R$" nunca fica sozinho no fim de uma linha
    return "R$ " +f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v, sinal=False):
    txt = f"{abs(v) * 100:.1f}".replace(".", ",") + "%"
    if sinal:
        return ("+" if v >= 0 else "-") + txt
    return txt


def destaques(r):
    """Frases prontas pro gerente — o que ele leria primeiro no relatório."""
    frases = []
    faltando = [l["loja"] for l in r["lojas"] if not l["enviou"]]
    if faltando:
        frases.append(f"Atenção: sem dados utilizáveis de {', '.join(faltando)} — o total do mês está incompleto.")
    t = r["total"]
    frase = f"Faturamento líquido de {brl(t['liquido'])}"
    if t["variacao"] is not None:
        frase += f" ({pct(t['variacao'], sinal=True)} sobre {r['mes_anterior_extenso']})"
    if t["pct_meta"] is not None:
        frase += f", {pct(t['pct_meta'])} da meta consolidada"
    frases.append(frase + ".")

    com_meta = [l for l in r["lojas"] if l["pct_meta"] is not None and l["enviou"]]
    batidas = [l["loja"] for l in com_meta if l["pct_meta"] >= 1]
    if batidas:
        frases.append(f"Bateu a meta: {', '.join(batidas)}.")
    if com_meta:
        pior = min(com_meta, key=lambda l: l["pct_meta"])
        if pior["pct_meta"] < 1:
            falta = pior["meta"] - pior["liquido"]
            frases.append(f"Mais distante da meta: {pior['loja']}, com {pct(pior['pct_meta'])} "
                          f"(faltaram {brl(falta)}).")
    if r["categorias"] and t["liquido"] > 0:
        c = r["categorias"][0]
        frases.append(f"Categoria que mais faturou: {c['categoria']} ({pct(c['participacao'])} do total).")
    return frases
