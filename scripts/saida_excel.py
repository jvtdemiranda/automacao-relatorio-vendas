"""Relatório em Excel: resumo, vendas por dia, itens consolidados e qualidade dos dados."""

import io
import re
import zipfile
from datetime import datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

VERDE, VERMELHO, CINZA = "1E7F4F", "B3261E", "5F6368"
CABECALHO = PatternFill("solid", start_color="1F4E5A", end_color="1F4E5A")
BRL = '"R$" #,##0.00'
PCT = "0.0%"
DATA_FIXA = datetime(2026, 1, 1)


def texto(ws, linha, coluna, valor, **estilo):
    """
    Escreve texto que veio das planilhas das lojas sempre como texto.
    Sem isso, um produto cadastrado como "=HYPERLINK(...)" viraria uma
    fórmula de verdade no Excel do gerente (injeção de fórmula).
    """
    c = ws.cell(row=linha, column=coluna, value=None if valor is None else str(valor))
    c.data_type = "s"
    for k, v in estilo.items():
        setattr(c, k, v)
    return c


def cabecalho(ws, linha, nomes):
    for j, nome in enumerate(nomes, start=1):
        c = ws.cell(row=linha, column=j, value=nome)
        c.fill = CABECALHO
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center")


def larguras(ws, valores):
    for j, w in enumerate(valores, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w


def aba_resumo(wb, r):
    ws = wb.active
    ws.title = "Resumo"
    ws["A1"] = f"Relatório de vendas · {r['mes_extenso']}"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A2"] = "Farmácia Vitalis Manipulação · consolidado das 3 lojas"
    ws["A2"].font = Font(italic=True, color=CINZA)

    linha = 4
    for frase in r["destaques"]:
        ws.cell(row=linha, column=1, value="• " + frase)
        linha += 1

    linha += 1
    cabecalho(ws, linha, ["Loja", "Faturamento líquido", "Meta", "% da meta",
                          f"Variação vs {r['mes_anterior_extenso']}", "Vendas", "Ticket médio", "Devoluções"])
    inicio = linha + 1
    for l in r["lojas"] + [dict(r["total"], loja="Total")]:
        linha += 1
        ws.cell(row=linha, column=1, value=l["loja"]).font = Font(bold=l["loja"] == "Total")
        for col, chave, fmt in [(2, "liquido", BRL), (3, "meta", BRL), (4, "pct_meta", PCT), (5, "variacao", PCT),
                                (6, "vendas", "0"), (7, "ticket_medio", BRL), (8, "devolucoes", BRL)]:
            c = ws.cell(row=linha, column=col, value=l.get(chave))
            c.number_format = fmt
            if l["loja"] == "Total":
                c.font = Font(bold=True)
    for coluna, limite in [("D", 1), ("E", 0)]:
        faixa = f"{coluna}{inicio}:{coluna}{linha}"
        ws.conditional_formatting.add(faixa, CellIsRule(operator="lessThan", formula=[str(limite)], font=Font(color=VERMELHO, bold=True)))
        ws.conditional_formatting.add(faixa, CellIsRule(operator="greaterThanOrEqual", formula=[str(limite)], font=Font(color=VERDE, bold=True)))

    linha += 2
    cabecalho(ws, linha, ["Categoria", "Faturamento líquido", "Participação"])
    for c in r["categorias"]:
        linha += 1
        ws.cell(row=linha, column=1, value=c["categoria"])
        ws.cell(row=linha, column=2, value=c["liquido"]).number_format = BRL
        ws.cell(row=linha, column=3, value=c["participacao"]).number_format = PCT
    larguras(ws, [26, 20, 14, 12, 22, 10, 14, 14])


def aba_diario(wb, r):
    ws = wb.create_sheet("Vendas por dia")
    cabecalho(ws, 1, ["Data", "Faturamento líquido"])
    for i, d in enumerate(r["diario"], start=2):
        ws.cell(row=i, column=1, value=datetime.fromisoformat(d["data"])).number_format = "dd/mm/yyyy"
        ws.cell(row=i, column=2, value=d["liquido"]).number_format = BRL
    larguras(ws, [14, 20])


def aba_itens(wb, dados):
    ws = wb.create_sheet("Itens consolidados")
    nomes = ["Data", "Loja", "Venda", "Categoria", "Produto", "Qtd", "Valor"]
    cabecalho(ws, 1, nomes)
    ordenado = dados.sort_values(["data", "loja", "venda"], kind="stable")
    for i, l in enumerate(ordenado.itertuples(index=False), start=2):
        ws.cell(row=i, column=1, value=datetime(l.data.year, l.data.month, l.data.day)).number_format = "dd/mm/yyyy"
        texto(ws, i, 2, l.loja)
        texto(ws, i, 3, l.venda)
        texto(ws, i, 4, l.categoria)
        texto(ws, i, 5, l.produto)
        ws.cell(row=i, column=6, value=int(l.qtd))
        ws.cell(row=i, column=7, value=float(l.valor)).number_format = BRL
    if len(ordenado):
        tabela = Table(displayName="itens", ref=f"A1:G{len(ordenado) + 1}")
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(tabela)
    ws.freeze_panes = "A2"
    larguras(ws, [12, 13, 11, 24, 36, 6, 12])


def aba_qualidade(wb, r):
    ws = wb.create_sheet("Qualidade dos dados")
    cabecalho(ws, 1, ["Loja", "Arquivo", "Linhas lidas", "Linhas usadas", "O que foi corrigido ou descartado"])
    for i, q in enumerate(r["qualidade"], start=2):
        texto(ws, i, 1, q["loja"])
        texto(ws, i, 2, q["arquivo"] or "planilha não recebida")
        ws.cell(row=i, column=3, value=q["linhas_lidas"])
        ws.cell(row=i, column=4, value=q["linhas_validas"])
        texto(ws, i, 5, "; ".join(q["problemas"]) or "nenhum problema encontrado")
    larguras(ws, [13, 32, 12, 13, 90])


def salvar(caminho, resumo, dados):
    wb = Workbook()
    aba_resumo(wb, resumo)
    aba_diario(wb, resumo)
    aba_itens(wb, dados)
    aba_qualidade(wb, resumo)
    wb.properties.creator = "automacao-relatorio-vendas"
    wb.properties.created = wb.properties.modified = DATA_FIXA

    # O .xlsx é um zip, e o openpyxl carimba a hora atual em cada arquivo
    # interno e na data de "modificado" (ignorando a que definimos acima).
    # Regravar com data fixa deixa o arquivo idêntico a cada execução: assim
    # o robô que publica os relatórios só commita quando algo mudou.
    buffer = io.BytesIO()
    wb.save(buffer)
    with zipfile.ZipFile(buffer) as origem, zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as destino:
        for item in origem.infolist():
            conteudo = origem.read(item.filename)
            if item.filename == "docProps/core.xml":
                conteudo = re.sub(rb"(<dcterms:modified[^>]*>)[^<]*", rb"\g<1>2026-01-01T00:00:00Z", conteudo)
            info = zipfile.ZipInfo(item.filename, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            destino.writestr(info, conteudo)
