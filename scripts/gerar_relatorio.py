"""
Gera o relatório mensal de vendas a partir das planilhas em entrada/AAAA-MM/.

    python gerar_relatorio.py            # todos os meses que têm pasta em entrada/
    python gerar_relatorio.py 2026-08    # só um mês

Para cada mês, escreve em saida/AAAA-MM/:
    relatorio.pdf   -> resumo executivo pra gerência
    relatorio.xlsx  -> mesmos números + itens consolidados, pra quem quer filtrar
    resumo.json     -> os números em formato de dados (é o que o CI confere)

E publica o mês mais recente em public/ (página + PDF + Excel pra baixar).
"""

import json
import os
import re
import shutil
import sys

import analise
import saida_excel
import saida_pdf
from leitura import ler_mes

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
ENTRADA = os.path.join(RAIZ, "entrada")


def meses_disponiveis():
    return sorted(p for p in os.listdir(ENTRADA) if re.fullmatch(r"\d{4}-\d{2}", p))


def gerar_mes(chave, metas):
    ano, mes = map(int, chave.split("-"))
    dados, registros = ler_mes(os.path.join(ENTRADA, chave), ano, mes)

    ano_ant, mes_ant = (ano, mes - 1) if mes > 1 else (ano - 1, 12)
    pasta_ant = os.path.join(ENTRADA, f"{ano_ant}-{mes_ant:02d}")
    dados_ant = ler_mes(pasta_ant, ano_ant, mes_ant)[0] if os.path.isdir(pasta_ant) else None

    resumo = analise.resumir(dados, registros, metas, ano, mes, dados_ant)

    saida = os.path.join(RAIZ, "saida", chave)
    os.makedirs(saida, exist_ok=True)
    saida_pdf.salvar(os.path.join(saida, "relatorio.pdf"), resumo)
    saida_excel.salvar(os.path.join(saida, "relatorio.xlsx"), resumo, dados)
    with open(os.path.join(saida, "resumo.json"), "w", encoding="utf-8") as f:
        json.dump(resumo, f, ensure_ascii=False, indent=2)
        f.write("\n")

    print(f"{chave}: {analise.brl(resumo['total']['liquido'])} líquidos · {len(dados)} itens")
    for q in resumo["qualidade"]:
        print(f"   {q['loja']:<11} {q['linhas_validas']:>4} de {q['linhas_lidas']:>4} linhas"
              + (f" · {'; '.join(q['problemas'])}" if q["problemas"] else ""))
    return resumo


def publicar(chave, resumo):
    """Página + arquivos do mês mais recente em public/ (o que a Vercel publica)."""
    publico = os.path.join(RAIZ, "public")
    os.makedirs(publico, exist_ok=True)
    for nome in ("relatorio.pdf", "relatorio.xlsx"):
        shutil.copyfile(os.path.join(RAIZ, "saida", chave, nome), os.path.join(publico, nome))

    dados_json = json.dumps(resumo, ensure_ascii=False, separators=(",", ":"))
    # Texto das planilhas (produtos, avisos) vai dentro de <script>: um
    # "</script>" ali fecharia a tag.
    dados_json = dados_json.replace("<", "\\u003c")
    with open(os.path.join(os.path.dirname(__file__), "pagina_template.html"), encoding="utf-8") as f:
        html = f.read().replace("/*__DADOS__*/", dados_json)
    with open(os.path.join(publico, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)


def main():
    disponiveis = meses_disponiveis()
    pedidos = sys.argv[1:] or disponiveis
    invalidos = [m for m in pedidos if m not in disponiveis]
    if invalidos:
        raise SystemExit(f"Sem pasta em entrada/ para: {', '.join(invalidos)}. Disponíveis: {', '.join(disponiveis)}")

    metas = analise.ler_metas(os.path.join(ENTRADA, "metas.csv"))
    resumos = {m: gerar_mes(m, metas) for m in pedidos}
    ultimo = max(resumos)
    if ultimo == disponiveis[-1]:
        publicar(ultimo, resumos[ultimo])
        print(f"Publicado em public/: {ultimo}")


if __name__ == "__main__":
    main()
